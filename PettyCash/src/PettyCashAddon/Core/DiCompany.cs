using System;
using SAPbobsCOM;

namespace PettyCashAddon.Core
{
    /// <summary>
    /// Connecte le DI API en réutilisant la session déjà ouverte par le
    /// client SAP B1 (via le "cookie" de contexte), sans redemander les
    /// identifiants à l'utilisateur. C'est le pattern standard pour tout
    /// add-on lancé depuis Add-On Administration.
    /// </summary>
    internal static class DiCompany
    {
        private static Company _company;

        public static Company Instance
        {
            get
            {
                if (_company == null)
                    throw new InvalidOperationException("DiCompany n'est pas encore connecté. Appeler Connect() au démarrage.");
                return _company;
            }
        }

        public static void Connect(SAPbouiCOM.Application sboApplication)
        {
            _company = new Company();

            string cookie = _company.GetContextCookie();
            string connectionContext = sboApplication.Company.GetConnectionContext(cookie);

            int rc = _company.SetSboLoginContext(connectionContext);
            if (rc != 0)
                throw new InvalidOperationException("SetSboLoginContext a échoué (code " + rc + ").");

            if (!_company.Connected)
            {
                int connectRc = _company.Connect();
                if (connectRc != 0)
                {
                    _company.GetLastError(out int errCode, out string errMsg);
                    throw new InvalidOperationException("Connexion DI API échouée : " + errCode + " - " + errMsg);
                }
            }
        }

        public static void Disconnect()
        {
            try
            {
                if (_company != null && _company.Connected)
                    _company.Disconnect();
            }
            catch
            {
                // arrêt de l'add-on : on ignore
            }
        }

        public static void ThrowIfError(int returnCode, string context)
        {
            if (returnCode != 0)
            {
                Instance.GetLastError(out int errCode, out string errMsg);
                throw new InvalidOperationException(context + " : [" + errCode + "] " + errMsg);
            }
        }
    }
}
