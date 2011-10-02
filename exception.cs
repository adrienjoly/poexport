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
 * File: Exception.cs
 *
 * Purpose: Managed representation of the IException interface
 *
 *
 * Notes: 
 *      This class does NOT subclass from System.Exception. It
 *      represents a cancelled or moved meeting.
 *
 */

namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    public class AppointmentException
    {
		private IntPtr m_pIAppointmentException;
		private Appointment m_appointment;
		private Application m_application;

        internal AppointmentException(Application application,
                                      Appointment appointment,
                            IntPtr pIAppointmentException)
        {
            m_application = application;
            m_appointment = appointment;
            m_pIAppointmentException =  pIAppointmentException;
        }

        public bool Deleted
        {
            get
            {
                int bDeleted = 0;
                PocketOutlook.CheckHRESULT(do_get_Deleted(m_pIAppointmentException, ref bDeleted));
                return bDeleted == 0 ? false : true;
            }
        } // Deleted

        public Appointment AppointmentItem
        {
            get
            {
                return m_appointment;
            }
        } // AppointmentItem

        public Application Application
        {
            get
            {
                return m_application;
            }
        } // Application

        [DllImport("PocketOutlook.dll", EntryPoint="IException_get_Deleted")]
        private static extern int do_get_Deleted(IntPtr pIAppointmentException, ref int rbDeleted);
    }
}
