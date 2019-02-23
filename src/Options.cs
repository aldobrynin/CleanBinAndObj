using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace CleanBinAndObj
{
    public class Options : DialogPage
    {
        // General
        [Category("General")]
        [DisplayName("Subdirectories to clean")]
        [Description("Subdirectories in project which will be cleaned")]
        [DefaultValue(true)]
        public string[] TargetSubdirectories { get; set; } = { "bin", "obj"};
}
}