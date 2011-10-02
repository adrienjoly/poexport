/*---------------------------------------------------------------------
   Copyright (C) Microsoft Corporation.  All rights reserved.

  This source code is intended only as a supplement to Microsoft
  Development Tools and/or on-line documentation.  See these other
  materials for detailed information regarding Microsoft code samples.

  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
  PARTICULAR PURPOSE.
 -----------------------------------------------------------------------
 * File: Application.cs
 *
 * Purpose: Managed representiation of the IPOutlookApp class
 *
 * Notes:
 *      The Application object is the only object which can be "new"ed
 *      by external libraries. Other PocketOutlook objects are created
 *      by calling various methods on the Application object.
 *
 *
 */
namespace PocketOutlook
{

    using System;
    using System.Runtime.InteropServices;
    using System.Diagnostics;

    public class Application
    {
		// the pointer to the unmanaged COM object
		private IntPtr m_pIPOutlookApp;

       /*
        * Do NOT call CoInitializeEx in managed code. The execution engine
        * takes care of this already.
        */

		public Application()
        {
            m_pIPOutlookApp = new IntPtr(0);
            int hResult = UnmanagedConstructor(ref m_pIPOutlookApp);

            PocketOutlook.CheckHRESULT(hResult);
        }

        public void Dispose()
        {
            //
            // Since m_pIPOutlookApp is an unmanaged object, it must be
            // destroyed by hand.
            //
            PocketOutlook.ReleaseCOMPtr(m_pIPOutlookApp);
        }

        ~Application()
        {
            this.Dispose();
        }

        /*
         * Property methods.
         *
         */

        public String Version
        {
            /*
             * Notes:
             *
             *    Native POOM property methods for "string"-like 
             *    properties (Such as IPOutlookApp::get_Version(...)) 
             *    all use BSTRs. BSTRs have a slightly unusual layout;
             *    however, all that matters here is that the actual
             *    pointer returned points to the beginning of a null
             *    terminated unicode string. So we can safely use
             *    System.Runtime.InteropServices.Marshal.PtrToStringUni(IntPtr)
             *    to extract the string.
             *
             *    [ Look up "BSTR" in the MSDN index for more information ]
             *
             */
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Version(m_pIPOutlookApp, ref bz);

                String zVersion = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zVersion = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    /*
                     * Each call to the native accessor allocates a new
                     * BSTR, so it must be freed each time.
                     */
                    PocketOutlook.SysFreeString(bz);
                }
                return zVersion;
            }
        } // Version


        /*
         * In native C++, the type of CurrentCityIndex is a 4-byte
         * "long". In managed C#, that's just an "int"
         */
        public int CurrentCityIndex
        {
            get
            {
                int iIndex = 0;
                int hResult = do_get_CurrentCityIndex(m_pIPOutlookApp, ref iIndex);
                PocketOutlook.CheckHRESULT(hResult);
                //return LookupCityIndex(iIndex);
                Debug.Assert(false, "LookupCityIndex not implemented!");
                return iIndex;
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_CurrentCityIndex(m_pIPOutlookApp,
                                                                        value));
            }
        } // CurrentCityIndex


        public City HomeCity
        {
            get
            {
                IntPtr pICity = new IntPtr(0);
                int hResult = do_get_HomeCity(m_pIPOutlookApp, ref pICity);

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                }
                catch (Exception)
                {
                    // release the reference to the object on failure
                    PocketOutlook.ReleaseCOMPtr(pICity);
                    throw;
                }

                return new City(this, ref pICity);
            }

            set
            {
                // the City object wraps a native ICity, so calling the
                // appropriate native function requires grabbing that
                // pointer.
                IntPtr pICity = value.RawCityPtr;
                int hResult = do_put_HomeCity(m_pIPOutlookApp, pICity);

                PocketOutlook.CheckHRESULT(hResult);
            }
        } // HomeCity

        
        public City VisitingCity
        {
            get
            {
                IntPtr pICity = new IntPtr(0);
                int hResult = do_get_VisitingCity(m_pIPOutlookApp, ref pICity);

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                }
                catch (Exception)
                {
                    PocketOutlook.ReleaseCOMPtr(pICity);
                    throw;
                }

                return new City(this, ref pICity);
            }

            set
            {
                PocketOutlook.CheckHRESULT(do_put_VisitingCity(m_pIPOutlookApp,
                                                            value.RawCityPtr));
            }
        } // VisitingCity

        
        public bool OutlookCompatible
        {
            get
            {
                bool bCompatible = false;
                int hResult = do_get_OutlookCompatible(m_pIPOutlookApp,
                                                       ref bCompatible);
                
                PocketOutlook.CheckHRESULT(hResult);

                return bCompatible;
            }
        } // OutlookCompatible


        public void Logon()
        {
            this.Logon(0);
        }

        public void Logon(int hWindowHandle)
        {
            PocketOutlook.CheckHRESULT(do_Logon(m_pIPOutlookApp,
                                                hWindowHandle));
        }

        public void Logoff()
        {
            PocketOutlook.CheckHRESULT(do_Logoff(m_pIPOutlookApp));
        }

        public Folder GetDefaultFolder(int tFolder)
        {
            IntPtr pIFolder = new IntPtr(0);
            int hResult = do_GetDefaultFolder(m_pIPOutlookApp,
                                            tFolder,
                                            ref pIFolder);
            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pIFolder);
                throw;
            }

            return new Folder(this, pIFolder);
        }


		/*
		 * Creation functions for the different Item types.
		 *
		 * Since a native enum can't readily be used in mananged code,
		 * and there are only 4 item types, we have seperate
		 * functions, rather than a single function
		 *      CreateItem(int itemType)
		 */

        public Appointment CreateAppointment()
		{
            IntPtr pIAppointment = new IntPtr(0);
            int hResult = do_CreateAppointment(m_pIPOutlookApp,
                                                ref pIAppointment);

            try
            {
				PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
      			PocketOutlook.ReleaseCOMPtr(pIAppointment);
                throw;
            }

            return new Appointment(this, ref pIAppointment);
        }
       
        public Contact CreateContact()
        {
            IntPtr pIContact = new IntPtr(0);
            int hResult = do_CreateContact(m_pIPOutlookApp,
                                            ref pIContact);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pIContact);
                throw;
            }

            return new Contact(this, ref pIContact);
        }
        
        public City CreateCity()
        {
            IntPtr pICity = new IntPtr(0);
            int hResult = do_CreateCity(m_pIPOutlookApp,
                                        ref pICity);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pICity);
                throw;
            }

            return new City(this, ref pICity);
        }
        
        public Task CreateTask()
        {
            IntPtr pITask = new IntPtr(0);
            int hResult = do_CreateTask(m_pIPOutlookApp,
                                        ref pITask);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pITask);
                throw;
            }

            return new Task(this, ref pITask);
        }

        public TimeZone GetTimeZoneFromIndex(int iIndex)
        {
            IntPtr pITimeZone = new IntPtr(0);
            int hResult = do_GetTimeZoneFromIndex(m_pIPOutlookApp,
                                                iIndex, ref pITimeZone);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pITimeZone);
                throw;
            }

            return new TimeZone(this, ref pITimeZone);
        }

        public void GetItemFromOid(long oid)
        {
            throw new NotSupportedException();
        }

		//
        // Imported functions
		//

        // Notes: 
        //
        // o The C# type "int" is equivalent to the C++ type
        //   "long"! Both types are 32 bit integers. If you mismatch
        //   types, you may assert inside the Execution Engine.
        // 
        // o A "IntPtr" appears in C++ as an intptr_t, which can be
        //   safely treated as any type of pointer (e.g. an
        //   IPOutlookApp*).
        //
        // o A "ref IntPtr" appears in C++ as a "intptr_t*", which can be
        //   safely treated as any type of double pointer (e.g. an
        //   IPOutlookApp**).
        // 
        // o A "String" appears as an "LPCTSTR", aka "const wchar_t*". See
        //   PocketOutlook.hpp for more.

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_SetupCOM") ]
        private static extern int do_SetupCOM();

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_Create") ]
        private static extern int UnmanagedConstructor(ref IntPtr rpApp);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_Logon") ]
        private static extern int do_Logon(IntPtr self, int hWindowHandle);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_Logoff") ]
        private static extern int do_Logoff(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_get_Version") ]
        private static extern int do_get_Version(IntPtr self,
                                                ref IntPtr rbzVersion);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_GetDefaultFolder") ]
        private static extern int do_GetDefaultFolder(IntPtr self,
                                                    int tFolder,
                                                    ref IntPtr rpFolder);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_CreateCity") ]
        private static extern int do_CreateCity(IntPtr self,
                                                ref IntPtr rpCity);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_CreateTask") ]
        private static extern int do_CreateTask(IntPtr self,
                                                ref IntPtr rpTask);
        
		[ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_CreateContact") ]
        private static extern int do_CreateContact(IntPtr self,
                                                ref IntPtr rpContact);
        
		[ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_CreateAppointment") ]
        private static extern int do_CreateAppointment(IntPtr self,
                                                ref IntPtr rpAppointment);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_GetItemFromOid") ]
        private static extern int do_GetItemFromOid(IntPtr self,
                                                    int Oid,
                                                    ref IntPtr rpItem);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_get_HomeCity") ]
        private static extern int do_get_HomeCity(IntPtr self,
                                                ref IntPtr rpCity);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_put_HomeCity") ]
        private static extern int do_put_HomeCity(IntPtr self,
                                                IntPtr pCity);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_get_VisitingCity") ]
        private static extern int do_get_VisitingCity(IntPtr self,
                                                    ref IntPtr rpCity);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_put_VisitingCity") ]
        private static extern int do_put_VisitingCity(IntPtr self,
                                                    IntPtr pCity);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_get_CurrentCityIndex") ]
        private static extern int do_get_CurrentCityIndex(IntPtr self,
                                                        ref int riIndex);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_put_CurrentCityIndex") ]
        private static extern int do_put_CurrentCityIndex(IntPtr self,
                                                        int iIndex);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_ReceiveFromInfrared") ]
        private static extern int do_ReceiveFromInfrared(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_get_OutlookCompatible") ]
        private static extern int do_get_OutlookCompatible(IntPtr self,
                                                            ref bool rbCompatible);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_GetTimeZoneFromIndex") ]
        private static extern int do_GetTimeZoneFromIndex(IntPtr self,
                                                        int cTimeZone,
                                                        ref IntPtr rpTimeZone);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_GetTimeZoneInformationFromIndex") ]
        private static extern int do_GetTimeZoneInformationFromIndex(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_SysFreeString") ]
        private static extern int do_SysFreeString(IntPtr self, IntPtr bz);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_VariantTimeToSystemTime") ]
        private static extern int do_VariantTimeToSystemTime(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="IPOutlookApp_SystemTimeToVariantTime") ]
        private static extern int do_SystemTimeToVariantTime(IntPtr self);


    } // class Application
} // namespace PocketOutlook
