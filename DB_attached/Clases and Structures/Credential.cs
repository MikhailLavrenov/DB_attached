using System;
using System.Xml.Serialization;
using System.Security.Cryptography;
using System.Text;

namespace DB_attached
{
    [Serializable]
    public class Credential    //структура учетных данных
    {
        [XmlIgnoreAttribute] public string Login { get; private set; }
        [XmlIgnoreAttribute] public string Password { get; private set; }
        public string ProtectedLogin { get; set; }
        public string ProtectedPassword { get; set; }
        public int Requests { get; set; }
        private string Encrypt(string text, int encryptionLevel)   //шифрует входной текст в соотвествии с заданным методом
        {
            if (encryptionLevel == 2)
                return text;

            byte[] byteText = Encoding.Default.GetBytes(text);

            DataProtectionScope scope;
            if (encryptionLevel == 0)
                scope = DataProtectionScope.CurrentUser;
            else
                scope = DataProtectionScope.LocalMachine;

            return Convert.ToBase64String(ProtectedData.Protect(byteText, null, scope));
        }
        private string Decrypt(string text, int encryptionLevel)   //расшифровывает входной текст в соотвествии с заданным методом
        {
            if (encryptionLevel == 2)
                return text;

            byte[] byteText = Convert.FromBase64String(text);

            DataProtectionScope scope;
            if (encryptionLevel == 0)
                scope = DataProtectionScope.CurrentUser;
            else
                scope = DataProtectionScope.LocalMachine;

            try
            {
                return Encoding.Default.GetString(ProtectedData.Unprotect(byteText, null, scope));
            }
            catch (Exception)
            {
                return "";
            }
        }
        public void SetLogin(string login, int encryptionLevel)   //устанавливает логин и шифрованный логин
        {
            this.Login = login;
            ProtectedLogin = Encrypt(login, encryptionLevel);
        }
        public void SetPassword(string password, int encryptionLevel)   //устанавливает пароль и зашифрованный пароль
        {
            this.Password = password;
            ProtectedPassword = Encrypt(password, encryptionLevel);
        }
        public void GenerateDecryptedCredential(int encryptionLevel)   //расшифровывает логин и пароль
        {
            this.Login = Decrypt(this.ProtectedLogin, encryptionLevel);
            this.Password = Decrypt(this.ProtectedPassword, encryptionLevel);
        }
        public Credential Copy()   //копирование класса
        {
            Credential cred = new Credential();
            cred.Login = string.Copy(this.Login);
            cred.Password = string.Copy(this.Password);
            cred.ProtectedLogin = string.Copy(this.ProtectedLogin);
            cred.ProtectedPassword = string.Copy(this.ProtectedPassword);
            cred.Requests = this.Requests;
            return cred;
        }
    }
}
