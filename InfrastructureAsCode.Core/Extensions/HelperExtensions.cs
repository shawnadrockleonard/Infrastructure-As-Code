using InfrastructureAsCode.Core.Enums;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace InfrastructureAsCode.Core.Extensions
{
    /// <summary>
    /// Provides common helper extensions for Office365 interactions
    /// </summary>
    public static class HelperExtensions
    {
        /// <summary>
        /// represents an additional Salt
        /// </summary>
        private static byte[] Entropy { get; set; }

        /// <summary>
        /// Contains a regular expression that are not valid for Windows or SharePoint storage
        /// </summary>
        private static string EscapedRegExpression { get; set; }

        /// <summary>
        /// Contains a regular expression that are not valid for Windows/SharePoint storage
        /// </summary>
        private static string EscapedPathRegExpression { get; set; }

        /// <summary>
        /// Represents invalid Hex Codes for invalid quickr formatting
        /// </summary>
        private static string EscapedHexExpression { get; set; }

        /// <summary>
        /// Compiled regular expression for performance.
        /// </summary>
        static Regex _htmlRegex = new Regex("<.*?>", RegexOptions.Compiled);

        /// <summary>
        /// Initialize the local variables
        /// </summary>
        static HelperExtensions()
        {
            Entropy = System.Text.Encoding.Unicode.GetBytes("PoSH_Automation");

            // clean filename of invalid characters
            // setup the characters that the file system does not like
            var invalidChars = System.IO.Path.GetInvalidFileNameChars().ToList();
            invalidChars.Add('–'); //adding hard dash as winzip doesn't like it
            invalidChars.AddRange(new char[] { '#', '%', '&', '+', ':' }); //adding sharepoint online sync characters
            EscapedRegExpression = string.Format("[{0}]", Regex.Escape(string.Join("", invalidChars)));

            var invalidPathChars = System.IO.Path.GetInvalidPathChars().ToList();
            invalidPathChars.Add('–'); //adding hard dash as winzip doesn't like it
            invalidPathChars.AddRange(new char[] { '#', '%', ':' }); //adding sharepoint online sync characters
            EscapedPathRegExpression = string.Format("[{0}]", Regex.Escape(string.Join("", invalidPathChars)));

            EscapedHexExpression = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
        }

        /// <summary>
        /// removes FBA user identity markup
        /// </summary>
        /// <param name="_user"></param>
        /// <returns></returns>
        public static string CleanUserString(this string _user)
        {
            string _cleanedUser = "";
            string[] _tmp = _user.Split(new char[] { '|' });
            string[] _tmp2 = _tmp.Last().Split(new char[] { '#' });
            _cleanedUser = _tmp2.Last();
            return _cleanedUser;
        }

        /// <summary>
        /// Read plain text password into Secure Credentials
        /// </summary>
        /// <param name="_username"></param>
        /// <param name="_password"></param>
        /// <returns></returns>
        public static SharePointOnlineCredentials GetCredentials(string _username, string _password)
        {
            SecureString passWord = new SecureString();
            foreach (char c in _password.ToCharArray()) passWord.AppendChar(c);
            var siteCredentials = new SharePointOnlineCredentials(_username, passWord);
            return siteCredentials;
        }

        public static string ConvertFromSecureString(this System.Security.SecureString input)
        {
            byte[] encryptedData = ProtectedData.Protect(Encoding.Unicode.GetBytes(ToInsecureString(input)), Entropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(encryptedData);
        }

        public static SecureString ConvertToSecureString(this string encryptedData)
        {
            try
            {
                byte[] decryptedData = ProtectedData.Unprotect(Convert.FromBase64String(encryptedData), Entropy, DataProtectionScope.CurrentUser);
                return ToSecureString(Encoding.Unicode.GetString(decryptedData));
            }
            catch (Exception ex)
            {
                var msg = string.Format("Exception: {0}", ex.Message);
                return new SecureString();
            }
        }

        public static SecureString ToSecureString(this string input)
        {
            SecureString secure = new SecureString();
            foreach (char c in input)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        public static string ToInsecureString(this SecureString input)
        {
            string returnValue = string.Empty;
            IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
            try
            {
                returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr);
            }
            return returnValue;
        }

        /// <summary>
        /// Uses regular expression to remove invalid characters for files
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="replaceValue"></param>
        /// <returns></returns>
        public static string GetCleanFileName(this string fileName, string replaceValue = "")
        {
            var newFileName = fileName;
            if (Regex.IsMatch(fileName, EscapedRegExpression))
            {
                newFileName = Regex.Replace(fileName, EscapedRegExpression, replaceValue, RegexOptions.IgnoreCase, new TimeSpan(10000));
            }

            var invalidChars = new char[] { '[', ']' };

            newFileName = newFileName.Replace(invalidChars, replaceValue);


            return newFileName;
        }

        /// <summary>
        /// Similiar to Regex and will replace the specific separators with the <paramref name="newValue"/>
        /// </summary>
        /// <param name="oldValue"></param>
        /// <param name="separators"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public static string Replace(this string oldValue, char[] separators, string newValue)
        {
            string[] temp;

            temp = oldValue.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return String.Join(newValue, temp);
        }

        /// <summary>
        /// Retreive the folder or create the directory
        /// </summary>
        /// <param name="rootDir"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static string GetOrCreateDirectory(this System.IO.DirectoryInfo rootDir, string folderName)
        {
            var outputPathDir = string.Empty;
            var pathContent = System.IO.Path.Combine(rootDir.FullName, folderName);
            if (!System.IO.Directory.Exists(pathContent))
            {
                var outputInfoPathDir = System.IO.Directory.CreateDirectory(pathContent, rootDir.GetAccessControl());
                outputPathDir = outputInfoPathDir.FullName;
                System.Diagnostics.Trace.TraceInformation("Directory {0} created", outputPathDir);
            }
            else
            {
                var outputInfoPathDir = new System.IO.DirectoryInfo(pathContent);
                outputPathDir = outputInfoPathDir.FullName;
                System.Diagnostics.Trace.TraceInformation("Directory {0} found", outputPathDir);
            }

            return outputPathDir;
        }

        /// <summary>
        /// Retreive the folder or create the directory
        /// </summary>
        /// <param name="rootDir"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        public static string GetOrCreateDirectory(this string rootDir, string folderName)
        {
            var outputPathDir = string.Empty;
            var pathContent = System.IO.Path.Combine(rootDir, folderName);
            if (!System.IO.Directory.Exists(pathContent))
            {
                var outputInfoPathDir = System.IO.Directory.CreateDirectory(pathContent);
                outputPathDir = outputInfoPathDir.FullName;
                System.Console.WriteLine(string.Format("Directory {0} created", outputPathDir));
            }
            else
            {
                var outputInfoPathDir = new System.IO.DirectoryInfo(pathContent);
                outputPathDir = outputInfoPathDir.FullName;
                System.Console.WriteLine(string.Format("Directory {0} found", outputPathDir));
            }

            return outputPathDir;
        }

        /// <summary>
        /// Uses regular expression to remove invalid characters for directories
        /// </summary>
        /// <param name="directoryName"></param>
        /// <param name="replaceValue"></param>
        /// <returns></returns>
        public static string GetCleanDirectory(this string directoryName, string replaceValue = "")
        {
            var trimmedFolder = directoryName.Trim();
            // Remove invalid characters
            if (Regex.IsMatch(directoryName, EscapedPathRegExpression))
            {
                trimmedFolder = Regex.Replace(trimmedFolder, EscapedPathRegExpression, replaceValue, RegexOptions.IgnoreCase, new TimeSpan(10000));
            }
            return trimmedFolder;
        }

        /// <summary>
        /// Uses regular expression to remove hex characters from poorly encoded text
        /// </summary>
        /// <param name="content"></param>
        /// <param name="replaceValue"></param>
        /// <returns></returns>
        public static string GetCleanContent(this string content, string replaceValue = "")
        {
            var trimmedContent = content.Trim();
            // Remove invalid characters
            if (Regex.IsMatch(trimmedContent, EscapedHexExpression))
            {
                trimmedContent = Regex.Replace(trimmedContent, EscapedHexExpression, replaceValue, RegexOptions.Compiled);
            }
            return trimmedContent;
        }

        /// <summary>
        /// Remove HTML from string with Regex.
        /// </summary>
        public static string StripTagsRegex(this string source)
        {
            return Regex.Replace(source, "<.*?>", string.Empty);
        }

        /// <summary>
        /// Remove HTML from string with compiled Regex.
        /// </summary>
        public static string StripTagsRegexCompiled(this string source)
        {
            return _htmlRegex.Replace(source, string.Empty);
        }
    }
}
