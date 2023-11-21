using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ADDON_PARAFLU.Services
{
    public class Security
    {
        // This constant determines the number of iterations for the password bytes generation function.
        private const int DerivationIterations = 10000;

        /// <summary>
        /// encrypt a string.
        /// </summary>
        /// <param name="plainText">the string the will be encrypt</param>
        /// <param name="passPhrase"> the chipher used for the ecryption </param>
        /// <returns> the encrypted string </returns>
        public static string Encrypt(string plainText)
        {
            if (string.IsNullOrWhiteSpace(plainText))
                return string.Empty;

            try
            {
                byte[] bytesBuff = Encoding.Unicode.GetBytes(plainText);
                string encryptText;

                using (Aes aes = Aes.Create())
                {
                    var crypto = new Rfc2898DeriveBytes(plainText, 32, DerivationIterations);
                    aes.Key = crypto.GetBytes(32);
                    aes.IV = crypto.GetBytes(16);

                    using (var memoryStream = new MemoryStream())
                    {
                        using (var cStream = new CryptoStream(memoryStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                        {
                            cStream.Write(bytesBuff, 0, bytesBuff.Length);
                            cStream.Close();
                        }

                        var cipherTextBytes = aes.Key;
                        cipherTextBytes = cipherTextBytes.Concat(aes.IV).ToArray();
                        cipherTextBytes = cipherTextBytes.Concat(memoryStream.ToArray()).ToArray();
                        encryptText = Convert.ToBase64String(cipherTextBytes);
                    }
                }

                return encryptText;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Decrypt a string
        /// </summary>
        /// <param name="cipherText"> the encrypted string </param>
        /// <param name="passPhrase"> the chypher used to encrypt </param>
        /// <returns> the string as a plain text. </returns>
        public static string Decrypt(string cipherText)
        {
            if (string.IsNullOrWhiteSpace(cipherText))
                return string.Empty;

            try
            {
                string decriptText;
                var cipherTextBytesWithSaltAndIv = Convert.FromBase64String(cipherText);
                // Get the saltbytes by extracting the first 32 bytes from the supplied cipherText bytes.
                var keyValue = cipherTextBytesWithSaltAndIv.Take(32).ToArray();
                // Get the IV bytes by extracting the next 32 bytes from the supplied cipherText bytes.
                var vectorValue = cipherTextBytesWithSaltAndIv.Skip(32).Take(16).ToArray();
                var cipherTextBytes = cipherTextBytesWithSaltAndIv.Skip(48).Take(cipherTextBytesWithSaltAndIv.Length - 48).ToArray();

                cipherText = cipherText.Replace(" ", "+");
                byte[] bytesBuff = cipherTextBytes;

                using (Aes aes = Aes.Create())
                {
                    aes.Key = keyValue;
                    aes.IV = vectorValue;

                    using (var memoryStream = new MemoryStream())
                    {
                        using (var cStream = new CryptoStream(memoryStream, aes.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cStream.Write(bytesBuff, 0, bytesBuff.Length);
                            cStream.Close();
                        }
                        memoryStream.Close();
                        decriptText = Encoding.Unicode.GetString(memoryStream.ToArray());
                        aes.Clear();
                    }
                }

                return decriptText;
            }
            catch
            {
                return cipherText;
            }
        }
    }
}
