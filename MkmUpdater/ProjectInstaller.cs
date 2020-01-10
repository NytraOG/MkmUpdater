using System.ComponentModel;
using System.Configuration.Install;

namespace MkmUpdater
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }
    }
}