namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Excel COM constants organized by category.
/// These map to Excel's XlXxx enumeration values.
/// </summary>
/// <remarks>
/// Reference: https://docs.microsoft.com/en-us/office/vba/api/overview/excel
/// </remarks>
public static class ExcelConstants
{
    #region Format Condition Types (XlFormatConditionType)

    /// <summary>XlFormatConditionType.xlCellValue - Format based on cell value</summary>
    public const int XlCellValue = 1;

    /// <summary>XlFormatConditionType.xlExpression - Format based on expression</summary>
    public const int XlExpression = 2;

    #endregion

    #region Format Condition Operators (XlFormatConditionOperator)

    /// <summary>XlFormatConditionOperator.xlBetween</summary>
    public const int XlBetween = 1;

    /// <summary>XlFormatConditionOperator.xlNotBetween</summary>
    public const int XlNotBetween = 2;

    /// <summary>XlFormatConditionOperator.xlEqual</summary>
    public const int XlEqual = 3;

    /// <summary>XlFormatConditionOperator.xlNotEqual</summary>
    public const int XlNotEqual = 4;

    /// <summary>XlFormatConditionOperator.xlGreater</summary>
    public const int XlGreater = 5;

    /// <summary>XlFormatConditionOperator.xlLess</summary>
    public const int XlLess = 6;

    /// <summary>XlFormatConditionOperator.xlGreaterEqual</summary>
    public const int XlGreaterEqual = 7;

    /// <summary>XlFormatConditionOperator.xlLessEqual</summary>
    public const int XlLessEqual = 8;

    #endregion

    #region Border Edges (XlBordersIndex)

    /// <summary>XlBordersIndex.xlEdgeLeft</summary>
    public const int XlEdgeLeft = 7;

    /// <summary>XlBordersIndex.xlEdgeTop</summary>
    public const int XlEdgeTop = 8;

    /// <summary>XlBordersIndex.xlEdgeBottom</summary>
    public const int XlEdgeBottom = 9;

    /// <summary>XlBordersIndex.xlEdgeRight</summary>
    public const int XlEdgeRight = 10;

    /// <summary>XlBordersIndex.xlInsideVertical</summary>
    public const int XlInsideVertical = 11;

    /// <summary>XlBordersIndex.xlInsideHorizontal</summary>
    public const int XlInsideHorizontal = 12;

    /// <summary>XlBordersIndex.xlDiagonalDown</summary>
    public const int XlDiagonalDown = 5;

    /// <summary>XlBordersIndex.xlDiagonalUp</summary>
    public const int XlDiagonalUp = 6;

    #endregion

    #region Validation Types (XlDVType)

    /// <summary>XlDVType.xlValidateInputOnly - Any value allowed</summary>
    public const int XlValidateInputOnly = 0;

    /// <summary>XlDVType.xlValidateWholeNumber</summary>
    public const int XlValidateWholeNumber = 1;

    /// <summary>XlDVType.xlValidateDecimal</summary>
    public const int XlValidateDecimal = 2;

    /// <summary>XlDVType.xlValidateList</summary>
    public const int XlValidateList = 3;

    /// <summary>XlDVType.xlValidateDate</summary>
    public const int XlValidateDate = 4;

    /// <summary>XlDVType.xlValidateTime</summary>
    public const int XlValidateTime = 5;

    /// <summary>XlDVType.xlValidateTextLength</summary>
    public const int XlValidateTextLength = 6;

    /// <summary>XlDVType.xlValidateCustom</summary>
    public const int XlValidateCustom = 7;

    #endregion

    #region Validation Alert Styles (XlDVAlertStyle)

    /// <summary>XlDVAlertStyle.xlValidAlertStop</summary>
    public const int XlValidAlertStop = 1;

    /// <summary>XlDVAlertStyle.xlValidAlertWarning</summary>
    public const int XlValidAlertWarning = 2;

    /// <summary>XlDVAlertStyle.xlValidAlertInformation</summary>
    public const int XlValidAlertInformation = 3;

    #endregion

    #region List Source Types (XlListSourceType)

    /// <summary>XlListSourceType.xlSrcRange</summary>
    public const int XlSrcRange = 1;

    /// <summary>XlListSourceType.xlSrcExternal</summary>
    public const int XlSrcExternal = 2;

    /// <summary>XlListSourceType.xlSrcQuery</summary>
    public const int XlSrcQuery = 3;

    #endregion

    #region Table Totals Calculation (XlTotalsCalculation)

    /// <summary>XlTotalsCalculation.xlTotalsCalculationNone</summary>
    public const int XlTotalsCalculationNone = 0;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationSum</summary>
    public const int XlTotalsCalculationSum = 1;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationAverage</summary>
    public const int XlTotalsCalculationAverage = 2;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationCount</summary>
    public const int XlTotalsCalculationCount = 3;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationCountNums</summary>
    public const int XlTotalsCalculationCountNums = 4;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationMin</summary>
    public const int XlTotalsCalculationMin = 5;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationMax</summary>
    public const int XlTotalsCalculationMax = 6;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationStdDev</summary>
    public const int XlTotalsCalculationStdDev = 7;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationVar</summary>
    public const int XlTotalsCalculationVar = 8;

    /// <summary>XlTotalsCalculation.xlTotalsCalculationCustom</summary>
    public const int XlTotalsCalculationCustom = 9;

    #endregion

    #region VBA Component Types (vbext_ComponentType)

    /// <summary>vbext_ct_StdModule - Standard module</summary>
    public const int VbaStdModule = 1;

    /// <summary>vbext_ct_ClassModule - Class module</summary>
    public const int VbaClassModule = 2;

    /// <summary>vbext_ct_MSForm - UserForm</summary>
    public const int VbaMSForm = 3;

    /// <summary>vbext_ct_Document - Document module (ThisWorkbook, Sheet1, etc.)</summary>
    public const int VbaDocument = 100;

    #endregion

    #region Insert/Delete Cells (XlInsertFormatOrigin)

    /// <summary>XlInsertFormatOrigin.xlFormatFromLeftOrAbove</summary>
    public const int XlFormatFromLeftOrAbove = 0;

    /// <summary>XlInsertFormatOrigin.xlFormatFromRightOrBelow</summary>
    public const int XlFormatFromRightOrBelow = 1;

    #endregion

    #region Insert/Delete Shift (XlInsertShiftDirection)

    /// <summary>XlInsertShiftDirection.xlShiftDown</summary>
    public const int XlShiftDown = -4121;

    /// <summary>XlInsertShiftDirection.xlShiftToRight</summary>
    public const int XlShiftToRight = -4161;

    #endregion

    #region Delete Shift (XlDeleteShiftDirection)

    /// <summary>XlDeleteShiftDirection.xlShiftUp</summary>
    public const int XlShiftUp = -4162;

    /// <summary>XlDeleteShiftDirection.xlShiftToLeft</summary>
    public const int XlShiftToLeft = -4159;

    #endregion

    #region List Object Source Type (XlListObjectSourceType)

    /// <summary>XlListObjectSourceType.xlSrcRange</summary>
    public const int XlListObjectSrcRange = 1;

    #endregion

    #region Pivot Table Creation (XlPivotTableSourceType)

    /// <summary>XlPivotTableSourceType.xlDatabase</summary>
    public const int XlDatabase = 1;

    #endregion

    #region Pivot Table Version (XlPivotTableVersionList)

    /// <summary>XlPivotTableVersionList.xlPivotTableVersion14 (Excel 2010+)</summary>
    public const int XlPivotTableVersion14 = 4;

    #endregion

    #region Command Types (XlCmdType)

    /// <summary>XlCmdType.xlCmdSql</summary>
    public const int XlCmdSql = 1;

    /// <summary>XlCmdType.xlCmdTable</summary>
    public const int XlCmdTable = 2;

    /// <summary>XlCmdType.xlCmdDefault</summary>
    public const int XlCmdDefault = 4;

    /// <summary>XlCmdType.xlCmdExcel - Power Query command type</summary>
    public const int XlCmdExcel = 6;

    #endregion
}
