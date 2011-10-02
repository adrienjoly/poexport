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
 * File: City.cs
 *
 * Purpose: Managed representation of the ICity object
 *
 *
 * Notes:
 *
 */

namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    public class City : OutlookItem
    {
        // The constructor is internal to prevent objects other than the
        // Application class from using "new" to create an object of this
        // type.
        internal City(Application application,
                    ref IntPtr pICity) : base (application, ref pICity)
        {
        }

        public void Dispose()
        {
            PocketOutlook.ReleaseCOMPtr(this.RawItemPtr);
        }

        ~City()
        {
            this.Dispose();
        }

        /*
         * Longitude and Latitude
         *
         * Notes:
         *      These accessors return integers which are "100 times the
         *      decimal representation of degrees". For example,
         *      Atlanta, GA, with a longitude of 84.42 degrees west,
         *      would return the value -8442. See the MSDN entry for
         *      "ICity Property Methods" for more details.
         */

        // Longitude
        public int Longitude
        {
            get
            {
                int nLongitude = 0;
                int hResult = do_get_Longitude(this.RawItemPtr, ref nLongitude);

                PocketOutlook.CheckHRESULT(hResult);

                return nLongitude;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Longitude(this.RawItemPtr, value));
            }
        } // Longitude


        public int Latitude
        {
            get
            {
                int nLatitude = 0;
                int hResult = do_get_Latitude(this.RawItemPtr, ref nLatitude);
                PocketOutlook.CheckHRESULT(hResult);
                return nLatitude;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Latitude(this.RawItemPtr, value));
            }
        } // Latitude

        
        /*
         * Returns the index for this city's timezone in the
         * Applications table of timezones.
         */
        public int TimezoneIndex
        {
            get
            {
                int iIndex = 0;
                int hResult = do_get_TimezoneIndex(this.RawItemPtr, ref iIndex);
                PocketOutlook.CheckHRESULT(hResult);
                return iIndex;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_TimezoneIndex(this.RawItemPtr,
                                                                value));
            }
        }

        public String AirportCode
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_AirportCode(this.RawItemPtr, ref bz);
                String zAirportCode = null;
                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zAirportCode = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zAirportCode;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_AirportCode(this.RawItemPtr, value));
            }
        } // AirportCode


        public String CountryPhoneCode
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_CountryPhoneCode(this.RawItemPtr, ref bz);
                String zCountryPhoneCode = null;

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zCountryPhoneCode = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zCountryPhoneCode;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_CountryPhoneCode(this.RawItemPtr, value));
            }
        } // CountryPhoneCode


        public String AreaCode
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_AreaCode(this.RawItemPtr, ref bz);
                String zAreaCode = null;

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zAreaCode = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zAreaCode;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_AreaCode(this.RawItemPtr, value));
            }
        } // AreaCode

        public String Name
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Name(this.RawItemPtr, ref bz);
                String zName = null;

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zName = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zName;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Name(this.RawItemPtr, value));
            }
        } // Name

        public String Country
        {
            get
            {
                IntPtr bz = new IntPtr(0);
                int hResult = do_get_Country(this.RawItemPtr, ref bz);
                String zCountry = null;

                try
                {
                    PocketOutlook.CheckHRESULT(hResult);
                    zCountry = Marshal.PtrToStringUni(bz);
                }
                finally
                {
                    PocketOutlook.SysFreeString(bz);
                }

                return zCountry;
            }
            set
            {
                PocketOutlook.CheckHRESULT(do_put_Country(this.RawItemPtr, value));
            }
        } // Country

        public bool InROM
        {
            get
            {
                bool bInROM = false;
                int hResult = do_get_InROM(this.RawItemPtr, ref bInROM);
                PocketOutlook.CheckHRESULT(hResult);
                return bInROM;
            }
        }

        protected override void doSave()
        {
            PocketOutlook.CheckHRESULT(do_Save(this.RawItemPtr));
        }

        protected override void doDelete()
        {
            PocketOutlook.CheckHRESULT(do_Delete(this.RawItemPtr));
        }

        protected override OutlookItem doCopy()
        {
            IntPtr pICity = new IntPtr(0);
            int hResult = do_Copy(this.RawItemPtr, ref pICity);

            try
            {
                PocketOutlook.CheckHRESULT(hResult);
            }
            catch (Exception)
            {
                PocketOutlook.ReleaseCOMPtr(pICity);
            }

            return new City(this.Application, ref pICity);
        }

        internal IntPtr RawCityPtr
        {
            get
            {
                return this.RawItemPtr;
            }
        }

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_Longitude") ]
        private static extern int do_get_Longitude(IntPtr self,
                                                    ref int pnLongitude);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_Latitude") ]
        private static extern int do_get_Latitude(IntPtr self,
                                                ref int pnLongitude);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_TimezoneIndex") ]
        private static extern int do_get_TimezoneIndex(IntPtr self,
                                                        ref int pnIndex);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_AirportCode") ]
        private static extern int do_get_AirportCode(IntPtr self,
                                                    ref IntPtr rbzAirportCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_CountryPhoneCode") ]
        private static extern int do_get_CountryPhoneCode(IntPtr self,
                                                        ref IntPtr rbzCountryPhoneCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_AreaCode") ]
        private static extern int do_get_AreaCode(IntPtr self,
                                                ref IntPtr rbzAreaCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_Name") ]
        private static extern int do_get_Name(IntPtr self,
                                            ref IntPtr rbzName);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_Country") ]
        private static extern int do_get_Country(IntPtr self,
                                                ref IntPtr rbzCountry);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_get_InROM") ]
        private static extern int do_get_InROM(IntPtr self,
                                                ref bool rbInROM);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_Longitude") ]
        private static extern int do_put_Longitude(IntPtr self,
                                                    int nLongitude);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_Latitude") ]
        private static extern int do_put_Latitude(IntPtr self,
                                                int nLatitude);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_TimezoneIndex") ]
        private static extern int do_put_TimezoneIndex(IntPtr self,
                                                        int nIndex);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_AirportCode") ]
        private static extern int do_put_AirportCode(IntPtr self,
                                                    String zAirportCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_CountryPhoneCode") ]
        private static extern int do_put_CountryPhoneCode(IntPtr self,
                                                        String zCountryPhoneCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_AreaCode") ]
        private static extern int do_put_AreaCode(IntPtr self,
                                                string zAreaCode);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_Name") ]
        private static extern int do_put_Name(IntPtr self,
                                            String zName);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_put_Country") ]
        private static extern int do_put_Country(IntPtr self,
                                                string zCountry);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_Save") ]
        private static extern int do_Save(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_Delete") ]
        private static extern int do_Delete(IntPtr self);

        [ DllImport("PocketOutlook.dll", EntryPoint="ICity_Copy") ]
        private static extern int do_Copy(IntPtr self,
                                        ref IntPtr rpICity);
    } // class City


}
