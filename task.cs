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
  * File: Task.cs
 *
 * Purpose: managed representation of the ITask object
 *
 *
 * Notes:
 *
 */

namespace PocketOutlook
{
	using System;
	using System.Runtime.InteropServices;

    public class Task : OutlookItem
    {
        internal Task(Application application, ref IntPtr pITask) : base (application, ref pITask)
        {
        }


        public String Subject
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Subject(this.RawItemPtr, ref bz);
                String zSubject = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zSubject = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zSubject;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Subject(this.RawItemPtr, value));
            }
        } // Subject

        public String Categories
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Categories(this.RawItemPtr, ref bz);
                String zCategories = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zCategories = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zCategories;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Categories(this.RawItemPtr, value));
            }
        } // Categories

        public int Importance
        {
            get
            {
                int nImportance = 0;
                int hResult = do_get_Importance(this.RawItemPtr, ref nImportance);
                PocketOutlook.CheckHRESULT(hResult);
                return nImportance;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Importance(this.RawItemPtr, value));
            }
        } // Importance

        public bool Complete
        {
            get
            {
                int bComplete = 0;
                PocketOutlook.CheckHRESULT(do_get_Complete(this.RawItemPtr, ref bComplete));
                return bComplete == 0 ? false : true;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Complete(this.RawItemPtr, value ? 1 : 0));
            }
        }

        public bool IsRecurring
        {
            get
            {
                int bIsRecurring = 0;
                PocketOutlook.CheckHRESULT(do_get_IsRecurring(this.RawItemPtr, ref bIsRecurring));
                return bIsRecurring == 0 ? false : true;
            }
        } // IsRecurring

        public int Sensitivity
        {
            get
            {
                int nSensitivity = 0;
                PocketOutlook.CheckHRESULT(do_get_Sensitivity(this.RawItemPtr, ref nSensitivity));
                return nSensitivity;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Sensitivity(this.RawItemPtr, value));
            }
        }

        public bool TeamTask
        {
            get
            {
                int bTeamTask = 0;
                PocketOutlook.CheckHRESULT(do_get_TeamTask(this.RawItemPtr, ref bTeamTask));
                return bTeamTask == 0 ? false : true;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_TeamTask(this.RawItemPtr,
                                                        value ? 1 : 0));
            }
        } // TeamTask

        public bool ReminderSet
        {
            get
            {
                int bReminderSet = 0;
                PocketOutlook.CheckHRESULT(do_get_ReminderSet(this.RawItemPtr, ref bReminderSet));
                return bReminderSet == 0 ? false : true;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_ReminderSet(this.RawItemPtr, value ? 1 : 0));
            }
        } // ReminderSet

        public String ReminderSoundFile
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_ReminderSoundFile(this.RawItemPtr, ref bz);
                String zReminderSoundFile = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zReminderSoundFile = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zReminderSoundFile;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_ReminderSoundFile(this.RawItemPtr, value));
            }
        } // ReminderSoundFile

        public int ReminderOptions
        {
            get
            {
                int nReminderOptions = 0;
                PocketOutlook.CheckHRESULT(do_get_ReminderOptions(this.RawItemPtr,
                                                                ref nReminderOptions));
                return nReminderOptions;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_ReminderOptions(this.RawItemPtr,
                                                                value));
            }
        } // ReminderOptions

        public String Body
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Body(this.RawItemPtr, ref bz);
                String zBody = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zBody = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zBody;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Body(this.RawItemPtr, value));
            }
        } // Body


        protected override void doSave()
        {
			throw new NotSupportedException();
        }

        protected override void doDelete()
        {
			throw new NotSupportedException();
        }

        protected override OutlookItem doCopy()
        {
			throw new NotSupportedException();
        }

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_ClearRecurrencePattern") ]
        private static extern int do_ClearRecurrencePattern(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_GetRecurrencePattern") ]
        private static extern int do_GetRecurrencePattern(IntPtr self,
                                                        IntPtr rpIRecurrencePattern);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_IsRecurring") ]
        private static extern int do_get_IsRecurring(IntPtr self,
                                                    ref int bIsRecurring);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Subject") ]
        private static extern int do_get_Subject(IntPtr self,
                                                ref IntPtr rbzSubject);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Categories") ]
        private static extern int do_get_Categories(IntPtr self,
                                                    ref IntPtr rbzCategories);

        // date items; these are not currently supported
        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_StartDate") ]
        private static extern int do_get_StartDate(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_DueDate") ]
        private static extern int do_get_DueDate(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_DateCompleted") ]
        private static extern int do_get_DateCompleted(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Importance") ]
        private static extern int do_get_Importance(IntPtr self,
                                                    ref int nImportance);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Complete") ]
        private static extern int do_get_Complete(IntPtr self,
                                                ref int rbComplete);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Sensitivity") ]
        private static extern int do_get_Sensitivity(IntPtr self,
                                                    ref int rnSensitivity);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_TeamTask") ]
        private static extern int do_get_TeamTask(IntPtr self,
                                                ref int rbTeamTask);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Body") ]
        private static extern int do_get_Body(IntPtr self,
                                            ref IntPtr rzBody);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_ReminderSet") ]
        private static extern int do_get_ReminderSet(IntPtr self,
                                                    ref int rbReminderSet);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_ReminderSoundFile") ]
        private static extern int do_get_ReminderSoundFile(IntPtr self,
                                                        ref IntPtr rzReminderSoundFile);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_ReminderTime") ]
        private static extern int do_get_ReminderTime(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_ReminderOptions") ]
        private static extern int do_get_ReminderOptions(IntPtr self,
                                                        ref int rnReminderOptions);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Subject") ]
        private static extern int do_put_Subject(IntPtr self, String zBody);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Categories") ]
        private static extern int do_put_Categories(IntPtr self, String zCategories);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_StartDate") ]
        private static extern int do_put_StartDate(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_DueDate") ]
        private static extern int do_put_DueDate(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Importance") ]
        private static extern int do_put_Importance(IntPtr self, int nImportance);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Complete") ]
        private static extern int do_put_Complete(IntPtr self, int bComplete);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Sensitivity") ]
        private static extern int do_put_Sensitivity(IntPtr self, int nSensitivity);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_TeamTask") ]
        private static extern int do_put_TeamTask(IntPtr self, int bTeamTask);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_Body") ]
        private static extern int do_put_Body(IntPtr self, String zBody);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_ReminderSet") ]
        private static extern int do_put_ReminderSet(IntPtr self, int
                                                    bReminderSet);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_ReminderSoundFile") ]
        private static extern int do_put_ReminderSoundFile(IntPtr self,
                                                        String zReminderSoundFile);
  
		[ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_ReminderTime") ]
        private static extern int do_put_ReminderTime(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_ReminderOptions") ]
        private static extern int do_put_ReminderOptions(IntPtr self,
                                                        long nReminderOptions);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_Save") ]
        private static extern int do_Save(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_Delete") ]
        private static extern int do_Delete(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_SkipRecurrence") ]
        private static extern int do_SkipRecurrence(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_Copy") ]
        private static extern int do_Copy(IntPtr self, IntPtr rpItem);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_Display") ]
        private static extern int do_Display(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_Oid") ]
        private static extern int do_get_Oid(IntPtr self, ref int rnOid);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_put_BodyInk") ]
        private static extern int do_put_BodyInk(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ITask_get_BodyInk") ]
        private static extern int do_get_BodyInk(IntPtr self);
    }
}
