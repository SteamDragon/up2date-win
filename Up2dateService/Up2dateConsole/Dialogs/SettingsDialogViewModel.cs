﻿using System;
using System.Windows.Input;
using Up2dateConsole.Helpers;
using Up2dateConsole.ServiceReference;
using Up2dateConsole.ViewService;

namespace Up2dateConsole.Dialogs
{
    public class SettingsDialogViewModel : DialogViewModelBase
    {
        private readonly IViewService viewService;
        private readonly IWcfClientFactory wcfClientFactory;
        private string tokenUrl;
        private string dpsUrl;

        public SettingsDialogViewModel(IViewService viewService, IWcfClientFactory wcfClientFactory)
        {
            this.viewService = viewService ?? throw new ArgumentNullException(nameof(viewService));
            this.wcfClientFactory = wcfClientFactory ?? throw new ArgumentNullException(nameof(wcfClientFactory));

            IsInitialized = Initialize();

            OkCommand = new RelayCommand(ExecuteOk, CanOk);
        }

        public bool IsInitialized { get; }

        public ICommand OkCommand { get; }

        public string TokenUrl
        {
            get => tokenUrl;
            set
            {
                if (tokenUrl == value) return;
                tokenUrl = value;
                OnPropertyChanged();
            }
        }

        public string DpsUrl
        {
            get => dpsUrl;
            set
            {
                if (dpsUrl == value) return;
                dpsUrl = value;
                OnPropertyChanged();
            }
        }

        private bool CanOk(object obj)
        {
            return !string.IsNullOrWhiteSpace(TokenUrl) && !string.IsNullOrWhiteSpace(DpsUrl);
        }

        private void ExecuteOk(object obj)
        {
            IWcfService service = null;
            string error = string.Empty;
            try
            {
                service = wcfClientFactory.CreateClient();
                service.SetRequestCertificateUrl(TokenUrl);
                service.SetProvisioningUrl(DpsUrl);
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            finally
            {
                wcfClientFactory.CloseClient(service);
            }

            if (!string.IsNullOrEmpty(error))
            {
                viewService.ShowMessageBox(Texts.ServiceAccessError);
                Close(false);
            }

            Close(true);
        }

        private bool Initialize()
        {
            IWcfService service = null;
            string error = string.Empty;
            try
            {
                service = wcfClientFactory.CreateClient();
                TokenUrl = service.GetRequestCertificateUrl();
                DpsUrl = service.GetProvisioningUrl();
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            finally
            {
                wcfClientFactory.CloseClient(service);
            }

            if (!string.IsNullOrEmpty(error))
            {
                viewService.ShowMessageBox(Texts.ServiceAccessError);
                return false;
            }

            return true;
        }

    }
}