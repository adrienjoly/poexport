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
 * File: Appointment.cs
 *
 * Purpose: Managed representiation of the IAppointment class
 *
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

    public class Appointment : OutlookItem
    {
		private Application m_application;

        // The constructor is internal to prevent objects outside of the
        // PocketOutlook library from using "new" to create an object of
        // this type.
        internal Appointment(Application application,
                            ref IntPtr pIAppointment) :
            base(application, ref pIAppointment)
        {
			m_application = application;
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

		public String Location
		{
			set
			{
				PocketOutlook.CheckHRESULT(do_put_Location(this.RawItemPtr, value));
			}
		} // Subject
		

		protected override void doSave()
		{
			PocketOutlook.CheckHRESULT(do_Save(this.RawItemPtr));
		}
		
		public void Send()
		{
			PocketOutlook.CheckHRESULT(do_Send(this.RawItemPtr));
		}

		protected override OutlookItem doCopy()
        {
            throw new NotSupportedException();
        }

        protected override void doDelete()
        {
			throw new NotSupportedException();
        }

		public Recipients Recipients
		{
			get
			{
				IntPtr pRecipients = new IntPtr(0);
				long hResult = do_get_Recipients(this.RawItemPtr, ref pRecipients);
				try
				{
					PocketOutlook.CheckHRESULT( (int) hResult);
				}
				catch (Exception)
				{
					PocketOutlook.ReleaseCOMPtr(pRecipients);
					throw;
				}
				return new Recipients(m_application, pRecipients);
			}
		}
	
		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_get_Subject") ]
		private static extern int do_get_Subject(IntPtr self,
			ref IntPtr rbzSubject);

		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_put_Subject") ]
		private static extern int do_put_Subject(IntPtr self, String zBody);
		
		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_get_Body") ]
		private static extern int do_get_Body(IntPtr self,
			ref IntPtr rzBody);

		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_put_Body") ]
		private static extern int do_put_Body(IntPtr self, String zBody);
	
		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_put_Location") ]
		private static extern int do_put_Location(IntPtr self, String zBody);
	
		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_Save") ]
		private static extern int do_Save(IntPtr self);
	
		[ DllImport("PocketOutlook.dll", EntryPoint="IAppointment_Send") ]
		private static extern int do_Send(IntPtr self);
		
		
		[DllImport("PocketOutlook.dll", EntryPoint="IAppointment_get_Recipients")]
		private static extern int do_get_Recipients(IntPtr self, ref IntPtr pRecipients);
	}

}
