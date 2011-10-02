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
 * File: CContactItem.cs
 */

namespace PocketOutlook
{
	using System;

	public class CContactItem
	{
		public string m_szLastName;    
		public string m_szFirtName;
		public string m_szAddress;
	
		public override string ToString ()
		{
			return m_szFirtName + " " + m_szLastName + " (" + 
				m_szAddress + ")";
		}
	}
}
