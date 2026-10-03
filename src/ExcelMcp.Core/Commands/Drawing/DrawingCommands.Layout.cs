using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Drawing;

public sealed partial class DrawingCommands
{
    /// <inheritdoc />
    public DrawingObjectListResult GroupObjects(IExcelBatch batch, string sheetName, List<string> objectNames, string? groupName = null)
    {
        ValidateSelection(objectNames, 2);
        ValidateLayoutName(groupName);
        return WithSelection(batch, sheetName, objectNames, (shapes, selection, ct) =>
        {
            EnsureUniqueLayoutName(shapes, groupName, ct);
            Excel.Shape? group = null;
            try
            {
                group = selection.Group();
                ApplyName(group, groupName);
                return LayoutResult(batch, [ReadDrawingObject(group, sheetName, ct)]);
            }
            finally
            {
                ComUtilities.Release(ref group);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectListResult UngroupObject(IExcelBatch batch, string sheetName, string objectName)
    {
        ValidateSelection([objectName], 1);
        return WithSelection(batch, sheetName, [objectName], (_, selection, ct) =>
        {
            Excel.Shape? group = null;
            Excel.ShapeRange? members = null;
            try
            {
                group = selection.Item(1);
                if (ReadKind(group) != DrawingObjectKind.Group)
                    throw new ArgumentException($"Drawing object '{objectName}' is not a group.");
                members = group.Ungroup();
                return ReadLayoutRange(batch, members, sheetName, ct);
            }
            finally
            {
                ComUtilities.Release(ref members);
                ComUtilities.Release(ref group);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectListResult AlignObjects(IExcelBatch batch, string sheetName, List<string> objectNames, DrawingAlignment alignment)
    {
        ValidateSelection(objectNames, 2);
        if (!Enum.IsDefined(alignment)) throw new ArgumentOutOfRangeException(nameof(alignment));
        return WithSelection(batch, sheetName, objectNames, (_, selection, ct) =>
        {
            // PIA gap: Align takes Office.Core enums; this project intentionally has no office.dll reference.
            dynamic lateBoundSelection = (dynamic)(object)selection;
            lateBoundSelection.Align((int)alignment, 0);
            return ReadLayoutRange(batch, selection, sheetName, ct);
        });
    }

    /// <inheritdoc />
    public DrawingObjectListResult DistributeObjects(IExcelBatch batch, string sheetName, List<string> objectNames, DrawingDistribution distribution)
    {
        ValidateSelection(objectNames, 3);
        if (!Enum.IsDefined(distribution)) throw new ArgumentOutOfRangeException(nameof(distribution));
        return WithSelection(batch, sheetName, objectNames, (_, selection, ct) =>
        {
            // PIA gap: Distribute takes Office.Core enums without an office.dll reference.
            dynamic lateBoundSelection = (dynamic)(object)selection;
            lateBoundSelection.Distribute((int)distribution, 0);
            return ReadLayoutRange(batch, selection, sheetName, ct);
        });
    }

    /// <inheritdoc />
    public DrawingObjectListResult DuplicateObject(
        IExcelBatch batch, string sheetName, string objectName, string? newName = null, double offsetLeft = 10, double offsetTop = 10)
    {
        ValidateSelection([objectName], 1);
        ValidateLayoutName(newName);
        if (!double.IsFinite(offsetLeft)) throw new ArgumentOutOfRangeException(nameof(offsetLeft));
        if (!double.IsFinite(offsetTop)) throw new ArgumentOutOfRangeException(nameof(offsetTop));
        return WithSelection(batch, sheetName, [objectName], (shapes, selection, ct) =>
        {
            EnsureUniqueLayoutName(shapes, newName, ct);
            Excel.Shape? source = null;
            Excel.ShapeRange? copies = null;
            Excel.Shape? copy = null;
            try
            {
                source = selection.Item(1);
                ValidateEligibleLayoutObject(source, ct, rejectMacros: true);
                var left = Convert.ToSingle(source.Left + offsetLeft);
                var top = Convert.ToSingle(source.Top + offsetTop);
                if (!float.IsFinite(left) || !float.IsFinite(top))
                    throw new ArgumentException("Offsets must produce finite drawing positions within Excel's numeric range.");
                copies = selection.Duplicate();
                copy = copies.Item(1);
                ApplyName(copy, newName);
                copy.Left = left;
                copy.Top = top;
                return ReadLayoutRange(batch, copies, sheetName, ct);
            }
            finally
            {
                ComUtilities.Release(ref copy);
                ComUtilities.Release(ref copies);
                ComUtilities.Release(ref source);
            }
        });
    }

    /// <inheritdoc />
    public DrawingObjectListResult SetZOrder(IExcelBatch batch, string sheetName, string objectName, DrawingZOrder zOrder)
    {
        ValidateSelection([objectName], 1);
        if (!Enum.IsDefined(zOrder)) throw new ArgumentOutOfRangeException(nameof(zOrder));
        return WithSelection(batch, sheetName, [objectName], (_, selection, ct) =>
        {
            // PIA gap: ZOrder takes an Office.Core enum without an office.dll reference.
            dynamic lateBoundSelection = (dynamic)(object)selection;
            lateBoundSelection.ZOrder((int)zOrder);
            return ReadLayoutRange(batch, selection, sheetName, ct);
        });
    }

    private static void ValidateSelection(List<string> names, int minimum)
    {
        ArgumentNullException.ThrowIfNull(names);
        if (names.Count < minimum || names.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException($"Provide at least {minimum} nonempty top-level drawing object names.", nameof(names));
        if (names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Count)
            throw new ArgumentException("Drawing object names must be distinct.", nameof(names));
    }

    private static void ValidateLayoutName(string? name)
    {
        if (name != null && (string.IsNullOrWhiteSpace(name) || name.Length > 255))
            throw new ArgumentException("A supplied drawing name must be nonempty and at most 255 characters.", nameof(name));
    }

    private static void EnsureUniqueLayoutName(Excel.Shapes shapes, string? name, CancellationToken ct)
    {
        if (name == null) return;
        Excel.Shape? existing = null;
        try
        {
            existing = FindShape(shapes, name, ct);
            if (existing != null) throw new ArgumentException($"Drawing object '{name}' already exists on this worksheet.");
        }
        finally
        {
            ComUtilities.Release(ref existing);
        }
    }

    private static DrawingObjectListResult WithSelection(
        IExcelBatch batch, string sheetName, List<string> names,
        Func<Excel.Shapes, Excel.ShapeRange, CancellationToken, DrawingObjectListResult> action)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.ShapeRange? selection = null;
            try
            {
                sheet = GetSheet(ctx.Book, sheetName);
                if (sheet.ProtectDrawingObjects)
                    throw new InvalidOperationException("Worksheet drawing objects are protected. Unprotect the worksheet before changing drawing layout.");
                shapes = sheet.Shapes;
                foreach (var name in names)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Shape? shape = null;
                    try
                    {
                        shape = FindShape(shapes, name, ct)
                            ?? throw new ArgumentException($"Drawing object '{name}' not found on sheet '{sheetName}'. Use list-objects to discover top-level names.");
                        ValidateEligibleLayoutObject(shape, ct);
                    }
                    finally
                    {
                        ComUtilities.Release(ref shape);
                    }
                }
                selection = shapes.Range[names.Cast<object>().ToArray()];
                return action(shapes, selection, ct);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateEligibleLayoutObject(Excel.Shape shape, CancellationToken ct, bool rejectMacros = false)
    {
        ct.ThrowIfCancellationRequested();
        if (ReadKind(shape) == DrawingObjectKind.Other)
            throw new ArgumentException($"Drawing object '{shape.Name}' has an unsupported type. ActiveX/OLE, charts and unknown drawing types are excluded.");
        if (rejectMacros && !string.IsNullOrEmpty(shape.OnAction))
            throw new ArgumentException($"Drawing object '{shape.Name}' has a macro assignment and cannot be duplicated.");
        if (ReadKind(shape) != DrawingObjectKind.Group) return;
        Excel.GroupShapes? members = null;
        try
        {
            members = shape.GroupItems;
            for (var i = 1; i <= members.Count; i++)
            {
                Excel.Shape? child = null;
                try
                {
                    child = members.Item(i);
                    ValidateEligibleLayoutObject(child, ct, rejectMacros);
                }
                finally
                {
                    ComUtilities.Release(ref child);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref members);
        }
    }

    private static DrawingObjectListResult LayoutResult(IExcelBatch batch, List<DrawingObjectInfo> objects) =>
        new() { Success = true, FilePath = batch.WorkbookPath, DrawingObjects = objects };

    private static DrawingObjectListResult ReadLayoutRange(IExcelBatch batch, Excel.ShapeRange range, string sheetName, CancellationToken ct)
    {
        var objects = new List<DrawingObjectInfo>();
        for (var i = 1; i <= range.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.Shape? shape = null;
            try
            {
                shape = range.Item(i);
                objects.Add(ReadDrawingObject(shape, sheetName, ct));
            }
            finally
            {
                ComUtilities.Release(ref shape);
            }
        }
        return LayoutResult(batch, objects);
    }
}
