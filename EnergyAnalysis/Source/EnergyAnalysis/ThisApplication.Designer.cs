/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 29-02-2024
 * Time: 17:01
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace EnergyAnalysis {

    public sealed partial class ThisApplication : Autodesk.Revit.UI.Macros.ApplicationEntryPoint {
        
        public event System.EventHandler Startup;
        
        public event System.EventHandler Shutdown;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void OnStartup() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            base.FinishInitialization();
            this.OnStartup();
            this.InternalStartup();
            if ((this.Startup != null)) {
                this.Startup(this, System.EventArgs.Empty);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnShutdown() {
            if ((this.Shutdown != null)) {
                this.Shutdown(this, System.EventArgs.Empty);
            }
            base.OnShutdown();
        }
        
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override string PrimaryCookie {
            get {
                return "ThisApplication";
            }
        }
    }
}
