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
 * File: Exceptions.cs
 *
 * Purpose: Managed representation of the IExceptions interface
 *
 *
 * Notes: 
 *      This is a wrapper around a collection of Exceptions.
 *
 */

namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    public class Exceptions
    {
		private Application m_application;
		private Appointment m_appointment;
		private IntPtr m_pIExceptions;
		
		internal Exceptions(Application application, Appointment appointment, IntPtr pIExceptions)
        {
            m_application = application;
            m_appointment = appointment;
            m_pIExceptions =  pIExceptions;
        }

        public int Count
        {
            get
            {
                int nCount = 0;
                PocketOutlook.CheckHRESULT(do_get_Count(m_pIExceptions, ref nCount));
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

        public AppointmentException Item(int iIndex)
        {
            IntPtr pIException = new IntPtr(0);
            int hResult = do_Item(m_pIExceptions, iIndex, ref pIException);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pIException);
                throw;
            }

            return new AppointmentException(m_application, m_appointment, pIException);
        }

        [DllImport("PocketOutlook.dll", EntryPoint="IExceptions_get_Count")]
        private static extern int do_get_Count(IntPtr pIException, ref int rnCount);

        [DllImport("PocketOutlook.dll", EntryPoint="IExceptions_Item")]
        private static extern int do_Item(IntPtr pIException, int iIndex, ref IntPtr rpIException);
    }
}
