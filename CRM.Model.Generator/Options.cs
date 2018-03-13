using CommandLine;

namespace Crm.Model.Generator
{
    internal partial class Program
    {
        private class Options
        {
            [Option('n', "namespace", Required = false, HelpText = "Namespace for code files. If absent, app.config value will be used.")]
            public string DefaultNamespace { get; set; }

            [Option('b', "basetype", Required = false, HelpText = "Name for base type for models. If absent, app.config value will be used.")]
            public string EntityBaseType { get; set; }

            [Option('p', "path", Required = false, HelpText = "Directory path for generated code. Will be created Crm.Model.Data subdir. Subdir will be cleared if exists! If absent, app.config value will be used.")]
            public string TargetPath { get; set; }

            [Option('c', "connectionstring", Required = false, HelpText = "Connection string to crm. If absent, app.config value will be used.")]
            public string CrmConnection { get; set; }

        }

    }
}
