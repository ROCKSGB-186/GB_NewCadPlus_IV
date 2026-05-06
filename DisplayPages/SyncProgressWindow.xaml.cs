using System;
using System.Windows;

namespace GB_NewCadPlus_IV
{
    public partial class SyncProgressWindow : Window
    {
        public SyncProgressWindow()
        {
            InitializeComponent();
        }

        public event EventHandler? CancelRequested
        {
            add => SyncProgressView.CancelRequested += value;
            remove => SyncProgressView.CancelRequested -= value;
        }

        public void UpdateProgress(SyncProgressInfo progress)
        {
            SyncProgressView.UpdateProgress(progress);
        }
    }
}
