using System;
using System.Configuration.Install;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	[RunInstaller(true)]
	public class ComInstaller : Installer
	{
		public override void Install(System.Collections.IDictionary
			stateSaver)
		{
			base.Install(stateSaver);

			RegistrationServices regsrv = new RegistrationServices();
			if (!regsrv.RegisterAssembly(this.GetType().Assembly,
				AssemblyRegistrationFlags.SetCodeBase))
			{
				throw new InstallException("Failed To Register for COM");
			}
		}

		public override void Uninstall(System.Collections.IDictionary
			savedState)
		{
			base.Uninstall(savedState);

			RegistrationServices regsrv = new RegistrationServices();
			if (!regsrv.UnregisterAssembly(this.GetType().Assembly))
			{
				throw new InstallException("Failed To Unregister for COM");
			}
		}
	}
}

