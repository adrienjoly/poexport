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
 * File: Recipients.class
 *
 * Purpose: Mannged representation of the IRecipients object
 *
 *
 * Notes:
 *
 */
namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;
	using MsgBox = System.Windows.Forms.MessageBox;

    public class Recipients
    {
		private Application m_application;
		private IntPtr m_pIRecipients;

        internal Recipients(Application application, IntPtr pIRecipients)
        {
			m_application = application;
            m_pIRecipients =  pIRecipients;
        }

        public int Count
        {
            get
            {
                int nCount = 0;
                PocketOutlook.CheckHRESULT(do_get_Count(m_pIRecipients, ref nCount));
                return nCount;
            }
        }

        public Application Application
        {
            get
            {
                return m_application;
            }
        }

		public int GetValue() 
		{
			return 5;
		}

        public Recipient AddRecipient(String zName)
		{
      		IntPtr pIRecipient = new IntPtr(0);
            int hResult = do_Add(m_pIRecipients, zName, ref pIRecipient);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pIRecipient);
                throw;
            }

            return new Recipient(m_application, pIRecipient);
		}

        public Recipient Item(int iIndex)
        {
            IntPtr pIRecipient = new IntPtr(0);
            int hResult = do_Item(m_pIRecipients, iIndex, ref pIRecipient);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pIRecipient);
                throw;
            }

            return new Recipient(m_application, pIRecipient);
        }

        public void Remove(int iIndex)
        {
            PocketOutlook.CheckHRESULT(do_Remove(m_pIRecipients, iIndex));
        }

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipients_get_Count")]
        private static extern int do_get_Count(IntPtr pIRecipients, ref int rnCount);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipients_Add")]
        private static extern int do_Add(IntPtr pIRecipients, String zName, ref IntPtr rpIRecipient);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipients_Item")]
        private static extern int do_Item(IntPtr pIRecipients, int iIndex, ref IntPtr rpIRecipient);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecipients_Remove")]
        private static extern int do_Remove(IntPtr pIRecipients, int iIndex);
    } // class Recipients
}
