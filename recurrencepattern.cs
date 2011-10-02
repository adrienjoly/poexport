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
 * File: RecurrencePattern.cs
 *
 * Purpose: managed representation of the IRecurrencePattern object
 *
 *
 * Notes:
 *
 */

namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    public class RecurrencePattern
    {
		private Application m_application;
		private Appointment m_appointment;
		private IntPtr m_pIRecurrencePattern;

        internal RecurrencePattern(Application application,
                                   Appointment appointment,
                                   IntPtr pIRecurrencePattern)
        {
            m_application = application;
            m_appointment = appointment;
            m_pIRecurrencePattern =  pIRecurrencePattern;
        }

        public int RecurrenceType
        {
            get
            {
                int nRecurrenceType = 0;
                PocketOutlook.CheckHRESULT(do_get_RecurrenceType(m_pIRecurrencePattern, ref nRecurrenceType));
                return nRecurrenceType;
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_RecurrenceType(m_pIRecurrencePattern, value));
            }
        } // RecurrenceType

        // PatternStartDate
        
        // PatternEndDate

        public bool NoEndDate
        {
            get
            {
                int bNoEndDate = 0;
                PocketOutlook.CheckHRESULT(do_get_NoEndDate(m_pIRecurrencePattern, ref bNoEndDate));
                return bNoEndDate == 0 ? false : true;
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_NoEndDate(m_pIRecurrencePattern, value ? 1 : 0));
            }
        } // NoEndDate

        public int Occurrences
        {
            get
            {
                int nOccurrences = 0;
                PocketOutlook.CheckHRESULT(do_get_Occurrences(m_pIRecurrencePattern, ref nOccurrences));
                return nOccurrences;
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_Occurrences(m_pIRecurrencePattern, value));
            }
        } // Occurrences

        // StartTime

        public int Duration
        {
            get
            {
                int nDuration = 0;
                PocketOutlook.CheckHRESULT(do_get_Duration(m_pIRecurrencePattern, ref nDuration));
                return nDuration;
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_Duration(m_pIRecurrencePattern, value));
            }
        } // Duration

        // EndTime

        public Exceptions Exceptions
        {
            get
            {
                IntPtr pIExceptions = new IntPtr(0);
                int hResult = do_get_Exceptions(m_pIRecurrencePattern, ref pIExceptions);
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                }
                catch (Exception)
                {
                    PocketOutlook.ReleaseCOMPtr(pIExceptions);
                }

                return new Exceptions(m_application, m_appointment, pIExceptions);
            }
        }

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_get_RecurrenceType")]
        private static extern int do_get_RecurrenceType(IntPtr pIRecurrencePattern, ref int rnRecurrenceType);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_get_NoEndDate")]
        private static extern int do_get_NoEndDate(IntPtr pIRecurrencePattern, ref int rbNoEndDate);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_get_Occurrences")]
        private static extern int do_get_Occurrences(IntPtr pIRecurrencePattern, ref int rnRecurrences);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_get_Duration")]
        private static extern int do_get_Duration(IntPtr pIRecurrencePattern, ref int rnDuration);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_get_Exceptions")]
        private static extern int do_get_Exceptions(IntPtr pIRecurrencePattern, ref IntPtr rpIExceptions);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_put_RecurrenceType")]
        private static extern int do_put_RecurrenceType(IntPtr pIRecurrencePattern, int nRecurrenceType);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_put_NoEndDate")]
        private static extern int do_put_NoEndDate(IntPtr pIRecurrencePattern, int bNoEndDate);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_put_Occurrences")]
        private static extern int do_put_Occurrences(IntPtr pIRecurrencePattern, int nRecurrences);

        [DllImport("PocketOutlook.dll", EntryPoint="IRecurrencePattern_put_Duration")]
        private static extern int do_put_Duration(IntPtr pIRecurrencePattern, int nDuration);
    } // class RecurrencePattern
}
