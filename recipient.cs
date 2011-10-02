/*---------------------------------------------------------------------
   Copyright (C) Microsoft Corporation.  All rights reserved.

  This source code is intended only as a supplement to Microsoft
  Development Tools and/or on-line documentation.  See these other
  materials for detailed information regarding Microsoft code samples.

  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
  PARTICULAR PURPOSE.
 ----------------------------------------------------------------------- *
 * File: Recipient.cs
 *
 * Purpose: managed representation of the IRecipient object
 *
 *
 * Notes:
 *
 */
namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    public class Recipient
    {
		private Application m_application;
		private IntPtr m_pIRecipient;

		internal Recipient(Application application, IntPtr pIRecipient)
        {
            m_application = application;
            m_pIRecipient =  pIRecipient;
        }

        public String Address
        {
            get
            {
                String zAddress = null;
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Address(m_pIRecipient,
                                                ref bz);
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }
                return zAddress;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Address(m_pIRecipient, value));
            }
        }

        public String Name
        {
            get
            {
                String zName = null;
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Name(m_pIRecipient,
                                                ref bz);
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);

                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }
                return zName;
            }
        }
        
        [DllImport("PocketOutlook.dll", EntryPoint="IRecipient_get_Address")]
        private static extern int do_get_Address(IntPtr pIRecipient, ref IntPtr rbzAddress);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipient_get_Name")]
        private static extern int do_get_Name(IntPtr pIRecipient, ref IntPtr rbzName);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipient_put_Address")]
        private static extern int do_put_Address(IntPtr pIRecipient, String zAddress);
    } // class Recipient
}
