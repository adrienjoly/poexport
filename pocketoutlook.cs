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
 * File: PocketOutlook.cs
 *
 * Purpose: internal free functions used by all classes in the
 *			PocketOutlook library.
 *
 * Notes:
 *
 */

namespace PocketOutlook
{
    using System;
    using System.Runtime.InteropServices;

    internal class PocketOutlook
    {
        public static void CheckHRESULT(int hResult)
        {
            if (Failed(hResult) != 0)
            {
                throw new Exception();
            }
        }

        [ DllImport("PocketOutlook.dll", EntryPoint="PocketOutlook_IsFailure") ]
        public extern static uint Failed(int hResult);

        [ DllImport("PocketOutlook.dll", EntryPoint="PocketOutlook_ReleaseCOMPtr") ]
        public extern static uint ReleaseCOMPtr(IntPtr p);

        [ DllImport("PocketOutlook.dll", EntryPoint="BSTR_SysFreeString") ]
        public extern static void SysFreeString(IntPtr p);
    }
}
