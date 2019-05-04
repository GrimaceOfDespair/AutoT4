using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace BennorMcCarthy.AutoT4
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidAutoT4PkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(typeof(Options), Options.CategoryName, Options.PageName, 1000, 1001, false)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class AutoT4Package : AsyncPackage
    {
        private const int CommandId = 0x0100;
        private static readonly Guid CommandSet = new Guid("92c31ba2-5827-4779-b3ff-cf0fed43e50a");

        private DTE _dte;
        Tuple<_dispBuildEvents_OnBuildBeginEventHandler, _dispBuildEvents_OnBuildDoneEventHandler> _buildEvents;

        private Options Options
        {
            get { return (Options)GetDialogPage(typeof(Options)); }
        }

        protected async override System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            base.Initialize();

            _dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            if (_dte == null)
            {
                throw new NotImplementedException("AutoT4 must run within Visual Studio");
            }

            await RegisterT4Files(_dte);

            _buildEvents = RegisterEvents(_dte);

            var commandService = (IMenuCommandService)GetService(typeof(IMenuCommandService));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand((object sender, EventArgs e) => this.RunTemplates(_dte), menuCommandID);
            commandService.AddCommand(menuItem);
        }

        private async System.Threading.Tasks.Task RegisterT4Files(DTE dte)
        {
            var objectExtenders = await GetServiceAsync(typeof(ObjectExtenders)) as ObjectExtenders;
            var extenderProvider = new AutoT4ExtenderProvider(dte);
            objectExtenders.RegisterExtenderProvider(VSConstants.CATID.CSharpFileProperties_string, AutoT4ExtenderProvider.Name, extenderProvider);
            objectExtenders.RegisterExtenderProvider(VSConstants.CATID.VBFileProperties_string, AutoT4ExtenderProvider.Name, extenderProvider);
        }

        private Tuple<_dispBuildEvents_OnBuildBeginEventHandler, _dispBuildEvents_OnBuildDoneEventHandler> RegisterEvents(DTE dte)
        {
            _dispBuildEvents_OnBuildBeginEventHandler buildBegin = (scope, action) => OnBuildBegin(dte, scope);
            _dispBuildEvents_OnBuildDoneEventHandler buildDone = (scope, action) => OnBuildDone(dte, scope);

            dte.Events.BuildEvents.OnBuildBegin += buildBegin;
            dte.Events.BuildEvents.OnBuildDone += buildDone;

            return Tuple.Create(buildBegin, buildDone);
        }

        private void OnBuildBegin(DTE dte, vsBuildScope scope)
        {
            RunTemplates(dte, scope, RunOnBuild.BeforeBuild, Options.RunOnBuild == DefaultRunOnBuild.BeforeBuild);
        }

        private void OnBuildDone(DTE dte, vsBuildScope scope)
        {
            RunTemplates(dte, scope, RunOnBuild.AfterBuild, Options.RunOnBuild == DefaultRunOnBuild.AfterBuild);
        }

        private void RunTemplates(DTE dte, vsBuildScope scope, RunOnBuild buildEvent, bool runIfDefault)
        {
            dte.GetProjectsWithinBuildScope(scope)
                .FindT4ProjectItems()
                .ThatShouldRunOn(buildEvent, runIfDefault)
                .ToList()
                .ForEach(item => item.RunTemplate());
        }

        private void RunTemplates(DTE dte)
        {
            dte.GetProjectsWithinBuildScope(vsBuildScope.vsBuildScopeSolution)
                .FindT4ProjectItems()
                .ToList()
                .ForEach(item => item.RunTemplate());
        }

        protected override int QueryClose(out bool canClose)
        {
            int result = base.QueryClose(out canClose);
            if (canClose && _dte != null && _buildEvents != null)
            {
                _dte.Events.BuildEvents.OnBuildBegin -= _buildEvents.Item1;
                _dte.Events.BuildEvents.OnBuildDone -= _buildEvents.Item2;
                _buildEvents = null;
            }
            return result;
        }
    }
}
