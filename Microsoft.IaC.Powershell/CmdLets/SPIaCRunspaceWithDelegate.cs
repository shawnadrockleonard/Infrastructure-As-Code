using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;
using PCommand = System.Management.Automation.Runspaces;

namespace Microsoft.IaC.Powershell.CmdLets
{
    /// <summary>
    /// /Initializes a runspace to import modules and execute cmdlets
    /// </summary>
    public class SPIaCRunspaceWithDelegate : IDisposable
    {
        /// <summary>
        /// Execution stack
        /// </summary>
        /// <remarks>This must be disposed</remarks>
        internal Runspace psRunSpace { get; private set; }

        /// <summary>
        /// Initialize the PS cmdlet to connect
        /// </summary>
        internal PCommand.Command connectCommand { get; private set; }

        public SPIaCRunspaceWithDelegate()
        {

        }

        public SPIaCRunspaceWithDelegate(SPIaCConnection connection)
        {
            // Create Initial Session State for runspace.
            InitialSessionState initialSession = InitialSessionState.CreateDefault();
            initialSession.ImportPSModule(new[] { "MSOnline" });

            // Create credential object.
            var credential = connection.GetActiveCredentials();

            // Create command to connect office 365.
            connectCommand = new PCommand.Command("Connect-MsolService");
            connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));

            psRunSpace = RunspaceFactory.CreateRunspace(initialSession);
            // Open runspace.
            psRunSpace.Open();
        }


        public Collection<PSObject> ExecuteRunspace(PCommand.Command paramCommand, string exceptionMsg)
        {
            var results = new Collection<PSObject>();

            try
            {
                //Iterate through each command and execute it.
                foreach (var iteratedCommand in new PCommand.Command[] { connectCommand, paramCommand })
                {
                    var pipe = psRunSpace.CreatePipeline();
                    pipe.Commands.Add(iteratedCommand);

                    // Execute command and generate results and errors (if any).
                    results = pipe.Invoke();
                    var error = pipe.Error.ReadToEnd();

                    if (error.Count > 0 && iteratedCommand == paramCommand)
                    {
                        throw new Exception(exceptionMsg);
                    }
                    else if (results.Count > 0 && iteratedCommand == paramCommand)
                    {
                        return results;
                    }
                }
            }
            catch (Exception ex)
            {
                // TODO: Implement an appropriate Exception Stack and logging here
                throw new Exception(ex.Message);
            }

            return results;
        }


        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (psRunSpace != null)
                    {
                        var psState = psRunSpace.RunspaceStateInfo.State;
                        if (psState != RunspaceState.Closed  || psState == RunspaceState.Opened)
                        {
                            // Close the runspace.
                            psRunSpace.Close();
                        }
                        psRunSpace.Dispose();
                    }
                }

                disposedValue = true;
            }
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        #endregion


    }
}
