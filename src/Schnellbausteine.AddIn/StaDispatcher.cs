using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Schnellbausteine.AddIn
{
    internal sealed class StaDispatcher : IDisposable
    {
        private static readonly TimeSpan CallbackUnwindDelay = TimeSpan.FromMilliseconds(100);
        private readonly ManualResetEventSlim started = new ManualResetEventSlim(false);
        private readonly Thread thread;
        private SynchronizationContext context;
        private Exception startupException;
        private int disposed;

        public StaDispatcher()
        {
            Log.Trace("dispatcher", "thread.create");
            thread = new Thread(RunMessageLoop);
            thread.IsBackground = true;
            thread.Name = "Schnellbausteine UI";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void Post(Action action)
        {
            if (action == null) throw new ArgumentNullException("action");
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
            {
                throw new ObjectDisposedException("StaDispatcher");
            }
            if (!started.Wait(TimeSpan.FromSeconds(5)))
            {
                throw new InvalidOperationException("Der Schnellbausteine-UI-Thread konnte nicht gestartet werden.");
            }
            if (startupException != null)
            {
                throw new InvalidOperationException(
                    "Der Schnellbausteine-UI-Thread konnte nicht gestartet werden.", startupException);
            }

            Log.Trace("dispatcher", "action.queue");
            context.Post(delegate
            {
                Log.Trace("dispatcher", "action.before");
                try
                {
                    action();
                }
                finally
                {
                    Log.Trace("dispatcher", "action.after");
                }
            }, null);
        }

        public void PostAfterCallback(
            Task callbackBodyCompleted, Action action, Action<Exception> rejected)
        {
            if (callbackBodyCompleted == null) throw new ArgumentNullException("callbackBodyCompleted");
            if (action == null) throw new ArgumentNullException("action");
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
            {
                throw new ObjectDisposedException("StaDispatcher");
            }

            Log.Trace("dispatcher", "deferred.queue");
            ThreadPool.QueueUserWorkItem(delegate
            {
                Log.Trace("dispatcher", "deferred.wait-callback");
                callbackBodyCompleted.Wait();
                Log.Trace("dispatcher", "deferred.callback-complete");
                Thread.Sleep(CallbackUnwindDelay);
                Log.Trace("dispatcher", "deferred.delay-complete");
                if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0) return;
                try
                {
                    Post(action);
                }
                catch (Exception exception)
                {
                    if (!(exception is ObjectDisposedException) && rejected != null)
                    {
                        rejected(exception);
                    }
                }
            });
        }

        public bool Shutdown(TimeSpan timeout)
        {
            Log.Trace("dispatcher", "shutdown.request");
            if (Interlocked.Exchange(ref disposed, 1) != 0) return !thread.IsAlive;
            if (!started.Wait(TimeSpan.FromSeconds(1)) || context == null) return !thread.IsAlive;

            if (Thread.CurrentThread == thread)
            {
                CloseFormsAndExit();
                return true;
            }

            context.Post(delegate { CloseFormsAndExit(); }, null);
            bool stopped = thread.Join(timeout);
            Log.Trace("dispatcher", stopped ? "shutdown.complete" : "shutdown.timeout");
            if (stopped) started.Dispose();
            return stopped;
        }

        public void Dispose()
        {
            Shutdown(TimeSpan.FromSeconds(2));
        }

        private void RunMessageLoop()
        {
            try
            {
                Log.Trace("dispatcher", "thread.entry");
                SynchronizationContext.SetSynchronizationContext(
                    new WindowsFormsSynchronizationContext());
                context = SynchronizationContext.Current;
                started.Set();
                Log.Trace("dispatcher", "thread.ready");
                if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0) return;
                Application.Run();
                Log.Trace("dispatcher", "thread.exit");
            }
            catch (Exception exception)
            {
                startupException = exception;
                started.Set();
                Log.Error("STA-UI-Thread", exception);
            }
        }

        private static void CloseFormsAndExit()
        {
            while (Application.OpenForms.Count > 0)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            Application.ExitThread();
        }
    }
}
