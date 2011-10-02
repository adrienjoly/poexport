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
 * File: Folder.cs
 *
 * Purpose: Managed representation of the IFolder interface
 *
 * Copyright (c) Microsoft, 2001-2002
 *
 * Notes: 
 *      This is a wrapper around an ItemCollection
 *
 */

namespace PocketOutlook
{
	using System;
	using System.Runtime.InteropServices;

	public class Folder
	{
		private Application m_application;
		private IntPtr m_pIFolder;

		internal Folder(Application application, IntPtr pIFolder)
		{
			m_application = application;
			m_pIFolder = pIFolder;
		}

	    public ItemCollection Items
		{
			get
	        {
		        IntPtr pItemCollection = new IntPtr(0);
			    long hResult = do_get_Items(m_pIFolder, ref pItemCollection);
				try
	            {
		            PocketOutlook.CheckHRESULT( (int) hResult);
			    }
	            catch (Exception)
		        {
			        PocketOutlook.ReleaseCOMPtr(pItemCollection);
				    throw;
	            }
		        return new ItemCollection(m_application, this.DefaultItemType, ref pItemCollection);
	        }
		}

	    public int DefaultItemType
		{
	        get
		    {
			    int tItemType = 0;
				PocketOutlook.CheckHRESULT(do_get_DefaultItemType(m_pIFolder, ref tItemType));
	            return tItemType;
		    }
		}

	    public Application Application
		{
			get
	        {
		        return m_application;
			}
	    }

	    [DllImport("PocketOutlook.dll", EntryPoint="IFolder_get_Items")]
		private static extern int do_get_Items(IntPtr pIFolder, ref IntPtr pItemCollection);

	    [DllImport("PocketOutlook.dll", EntryPoint="IFolder_get_DefaultItemType")]
		private static extern int do_get_DefaultItemType(IntPtr pIFolder, ref int tItemType);
	} 
} 
