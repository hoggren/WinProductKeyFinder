using System;
using System.Diagnostics;
using Microsoft.Win32;

namespace WinProductKeyFinder {
    internal class Program {
        private static void Main() {
            Console.WriteLine("** Retrieves your Windows product key **\n");

            try {
                byte[] binaryProductKey;
                using (var keys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion")) {
                    if (keys == null)
                        throw new ApplicationException("Could not load product key from registry");

                    binaryProductKey = (byte[]) keys.GetValue("DigitalProductId");
                }

                var productKey = ConvertToKey(binaryProductKey);

                Console.WriteLine($"Decoded product key: {productKey}\n\nPress any key to exit...");
                Console.ReadKey();
            }
            catch (ApplicationException ex) {
                Console.WriteLine($"Registry loading error, exception message: {ex.Message}");
            }
        }

        /// <summary>
        ///     Decodes product key from registry.
        ///     Thanks to https://gallery.technet.microsoft.com/scriptcenter/How-to-backup-Windows-7367e9c5 for inspiration! :)
        /// </summary>
        /// <param name="binaryKey"></param>
        /// <returns></returns>
        private static string ConvertToKey(byte[] binaryKey) {
            const int keyOffset = 52;
            const string insert = "N";
            const string maps = "BCDFGHJKMPQRTVWXY2346789";

            // product key output
            var keyOutput = "";

            // Check if OS is Windows 8 
            var isWin8 = (byte) ((binaryKey[66] / 6) & 1);

            // set win8/win10 char
            binaryKey[66] = (byte) ((binaryKey[66] & 0xF7) | ((isWin8 & 2) * 4));

            // saves the last position in the dowhile below
            int last;
            var i = 24;
            do {
                var current = 0;
                var j = 14;
                do {
                    current = current * 256;
                    current = binaryKey[j + keyOffset] + current;
                    binaryKey[j + keyOffset] = (byte) (current / 24);
                    current = current % 24;
                    j = j - 1;
                } while (j >= 0);

                i = i - 1;
                keyOutput = Mid(maps, current + 1, 1) + keyOutput;
                last = current;
            } while (i >= 0);

            var keypart1 = Mid(keyOutput, 2, last);

            keyOutput = keyOutput.Substring(1).Replace(keypart1, keypart1 + insert);

            if (last == 0)
                keyOutput = insert + keyOutput;

            for (var k = 5; k < keyOutput.Length; k += 6)
                keyOutput = keyOutput.Insert(k, "-");

            return keyOutput;
        }

        // snippet to use VBscript function Mid
        private static string Mid(string param, int startIndex, int length) => param.Substring((param.Length + startIndex - 1) % param.Length, length);
    }
}