using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;

namespace CRM.Model.Generator.Core
{
    public static class Helpers
    {
        public static string ToLiteral(string input)
        {
            using (var writer = new StringWriter())
            {
                using (var provider = CodeDomProvider.CreateProvider("CSharp"))
                {
                    provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
                    return writer.ToString();
                }
            }
        }
    }
}