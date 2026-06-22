using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace WordToExeAddin
{
    public class LicenseService
    {
        private const string RegPath = "Software\\Aksa10xFaster";
        private const int TrialDays = 30;
        private const string SecretSalt = "AKSA10X_V2_2026";
        private static readonly DateTime BaseDate = new DateTime(2026, 1, 1);
        private const string CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public enum LicenseStatus
        {
            TrialActive,
            TrialExpired,
            Licensed,
            LicenseExpired
        }

        public LicenseStatus GetStatus()
        {
            string key = GetRegStr("LicenseKey", "");
            if (!string.IsNullOrEmpty(key))
            {
                string clean = key.Trim().Replace("-", "").Replace(" ", "").ToUpper();
                if (clean.Length == 25 && ChecksumValid(clean))
                {
                    string expiryStr = clean.Substring(0, 3);
                    int monthOffset = DecodeBase36(expiryStr);
                    if (monthOffset >= 100000)
                        return LicenseStatus.Licensed;

                    DateTime expiry = BaseDate.AddMonths(monthOffset);
                    if (expiry >= DateTime.Now.Date)
                        return LicenseStatus.Licensed;

                    return LicenseStatus.LicenseExpired;
                }
            }

            string trialStart = GetRegStr("TrialStart", "");
            if (string.IsNullOrEmpty(trialStart))
            {
                SetRegStr("TrialStart", DateTime.Now.ToString("yyyy-MM-dd"));
                return LicenseStatus.TrialActive;
            }

            DateTime start;
            if (DateTime.TryParse(trialStart, out start))
            {
                if ((DateTime.Now - start).TotalDays <= TrialDays)
                    return LicenseStatus.TrialActive;
            }

            return LicenseStatus.TrialExpired;
        }

        public int GetTrialDaysLeft()
        {
            string trialStart = GetRegStr("TrialStart", "");
            if (string.IsNullOrEmpty(trialStart))
                return TrialDays;

            DateTime start;
            if (DateTime.TryParse(trialStart, out start))
            {
                int remaining = TrialDays - (int)(DateTime.Now - start).TotalDays;
                return remaining > 0 ? remaining : 0;
            }
            return 0;
        }

        public string GetTrialStartDateString()
        {
            string trialStart = GetRegStr("TrialStart", "");
            if (string.IsNullOrEmpty(trialStart))
                return DateTime.Now.ToString("dd-MMM-yyyy");
            return trialStart;
        }

        public string GetTrialEndDateString()
        {
            string trialStart = GetRegStr("TrialStart", "");
            if (string.IsNullOrEmpty(trialStart))
                return DateTime.Now.AddDays(TrialDays).ToString("dd-MMM-yyyy");

            DateTime start;
            if (DateTime.TryParse(trialStart, out start))
                return start.AddDays(TrialDays).ToString("dd-MMM-yyyy");

            return DateTime.Now.AddDays(TrialDays).ToString("dd-MMM-yyyy");
        }

        public string GetLicenseExpiryDateString()
        {
            string key = GetRegStr("LicenseKey", "");
            if (string.IsNullOrEmpty(key))
                return null;

            string clean = key.Trim().Replace("-", "").Replace(" ", "").ToUpper();
            if (clean.Length != 25 || !ChecksumValid(clean))
                return null;

            string expiryStr = clean.Substring(0, 3);
            int monthOffset = DecodeBase36(expiryStr);
            if (monthOffset >= 100000)
                return "Lifetime";

            DateTime expiry = BaseDate.AddMonths(monthOffset);
            return expiry.ToString("dd-MMM-yyyy");
        }

        public string GetLicenseDurationLabel()
        {
            string key = GetRegStr("LicenseKey", "");
            if (string.IsNullOrEmpty(key))
                return null;

            string clean = key.Trim().Replace("-", "").Replace(" ", "").ToUpper();
            if (clean.Length != 25 || !ChecksumValid(clean))
                return null;

            char dur = clean[23];
            switch (dur)
            {
                case '6': return "6 Months";
                case '1': return "1 Year";
                case '2': return "2 Years";
                case 'L': return "Lifetime";
                default: return "Unknown";
            }
        }

        public bool Register(string licenseKey)
        {
            if (!ValidateKey(licenseKey))
                return false;

            SetRegStr("LicenseKey", licenseKey.Trim().ToUpper());
            return true;
        }

        public void Reset()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegPath))
            {
                key.DeleteValue("LicenseKey", false);
                key.DeleteValue("TrialStart", false);
            }
        }

        public static bool ValidateKey(string key)
        {
            if (string.IsNullOrEmpty(key))
                return false;

            string clean = key.Trim().Replace("-", "").Replace(" ", "").ToUpper();
            if (clean.Length != 25)
                return false;

            for (int i = 0; i < 25; i++)
            {
                char c = clean[i];
                if (!((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')))
                    return false;
            }

            if (!ChecksumValid(clean))
                return false;

            string expiryStr = clean.Substring(0, 3);
            int monthOffset = DecodeBase36(expiryStr);
            if (monthOffset < 0)
                return false;

            if (monthOffset >= 100000)
                return true;

            DateTime expiry = BaseDate.AddMonths(monthOffset);
            return expiry >= DateTime.Now.Date;
        }

        public static bool ChecksumValid(string clean)
        {
            int checksum = 0;
            for (int i = 0; i < 24; i++)
            {
                char c = clean[i];
                checksum += (c >= '0' && c <= '9') ? (c - '0') : (c - 'A' + 10);
            }

            int expected = checksum % 36;
            char lastChar = clean[24];
            int lastVal = (lastChar >= '0' && lastChar <= '9') ? (lastChar - '0') : (lastChar - 'A' + 10);
            return lastVal == expected;
        }

        public static string GenerateKey(string seed, int durationMonths)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(seed + SecretSalt + durationMonths));
                char[] result = new char[25];

                DateTime expiry;
                if (durationMonths >= 100000)
                {
                    expiry = DateTime.MaxValue;
                }
                else
                {
                    expiry = DateTime.Now.Date.AddMonths(durationMonths);
                }

                string expiryCode;
                if (durationMonths >= 100000)
                {
                    expiryCode = "ZZZ";
                }
                else
                {
                    int monthOffset = (expiry.Year - BaseDate.Year) * 12 + (expiry.Month - BaseDate.Month);
                    expiryCode = EncodeBase36(monthOffset, 3);
                }

                for (int i = 0; i < 3; i++)
                    result[i] = expiryCode[i];

                for (int i = 3; i < 23; i++)
                    result[i] = CharSet[hash[(i - 3) % hash.Length] % 36];

                char durChar;
                if (durationMonths >= 100000) durChar = 'L';
                else if (durationMonths >= 24) durChar = '2';
                else if (durationMonths >= 12) durChar = '1';
                else durChar = '6';
                result[23] = durChar;

                int checksum = 0;
                for (int i = 0; i < 24; i++)
                {
                    char c = result[i];
                    checksum += (c >= '0' && c <= '9') ? (c - '0') : (c - 'A' + 10);
                }
                result[24] = CharSet[checksum % 36];

                return new string(result, 0, 5) + "-" +
                       new string(result, 5, 5) + "-" +
                       new string(result, 10, 5) + "-" +
                       new string(result, 15, 5) + "-" +
                       new string(result, 20, 5);
            }
        }

        private static string EncodeBase36(int value, int minLen)
        {
            string chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string result = "";
            do
            {
                result = chars[value % 36] + result;
                value /= 36;
            } while (value > 0);
            while (result.Length < minLen)
                result = "0" + result;
            return result;
        }

        private static int DecodeBase36(string str)
        {
            int result = 0;
            for (int i = 0; i < str.Length; i++)
            {
                char c = str[i];
                int val = (c >= '0' && c <= '9') ? (c - '0') : (c - 'A' + 10);
                result = result * 36 + val;
            }
            return result;
        }

        private static string GetRegStr(string name, string def)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RegPath))
                return key != null ? (key.GetValue(name, def) as string ?? def) : def;
        }

        private static void SetRegStr(string name, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegPath))
                key.SetValue(name, value);
        }
    }
}
