using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;

namespace HMGSeis
{

    /// <summary>
    /// 
    /// 
    /// </summary>
    public class AutoCADConnector : IDisposable
    {
        private AcadApplication _application;
        private bool _initialized;
        private bool _disposed;
        public AutoCADConnector()
        {
            try
            {

                // Upon creation, attempt to retrieve running instance  
                _application = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application.21");
            }
            catch
            {
                try
                {
                    // Create an instance and set flag to indicate this  
                    _application = new AcadApplication();
                    _initialized = true;
                }
                catch
                {
                    throw;
                }
            }
        }

        // If the user doesn't call Dispose, the  
        // garbage collector will upon destruction  
        ~AutoCADConnector()
        {
            Dispose(false);
        }
        public AcadApplication Application
        {
            get
            {
                // Return our internal instance of AutoCAD  
                return _application;
            }
        }

        // This is the user-callable version of Dispose.  
        // It calls our internal version and removes the  
        // object from the garbage collector's queue. 
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        // This version of Dispose gets called by our  
        // destructor.  
        protected virtual void Dispose(bool disposing)
        {

            // If we created our AutoCAD instance, call its  
            // Quit method to avoid leaking memory.  
            if (!this._disposed && _initialized)
                _application.Quit();

            _disposed = true;
        }
    }

}


