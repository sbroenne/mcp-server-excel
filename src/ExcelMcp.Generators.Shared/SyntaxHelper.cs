using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Sbroenne.ExcelMcp.Generators.Common;

/// <summary>
/// Helper methods for syntax analysis shared between generators.
/// </summary>
public static class SyntaxHelper
{
    /// <summary>
    /// Checks if a syntax node is an interface with attributes.
    /// </summary>
    public static bool IsInterfaceWithAttributes(SyntaxNode node)
    {
        return node is InterfaceDeclarationSyntax ids && ids.AttributeLists.Count > 0;
    }

    /// <summary>
    /// Gets the interface declaration if it has [ServiceCategory] attribute.
    /// </summary>
    public static InterfaceDeclarationSyntax? GetServiceInterfaceOrNull(GeneratorSyntaxContext context)
    {
        var interfaceDeclaration = (InterfaceDeclarationSyntax)context.Node;

        // Check if it has ServiceCategory attribute
        foreach (var attributeList in interfaceDeclaration.AttributeLists)
        {
            foreach (var attribute in attributeList.Attributes)
            {
                var name = attribute.Name.ToString();
                if (name == "ServiceCategory" || name == "ServiceCategoryAttribute")
                {
                    return interfaceDeclaration;
                }
            }
        }

        return null;
    }
}
