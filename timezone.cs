/*
 *------------------------------------------------------------------------------
 *  <copyright from='1997' to='2001' company='Microsoft Corporation'>
 *   Copyright (c) Microsoft Corporation. All Rights Reserved.
 *
 *   This source code is intended only as a supplement to Microsoft
 *   Development Tools and/or on-line documentation.  See these other
 *   materials for detailed information regarding Microsoft code samples.
 *
 *   </copyright>
 *-------------------------------------------------------------------------------
 * File: TimeZone.cs
 *
 * Purpose: managed reprezentation of the ITimeZone object
 *
 *
 * Notes:
 *
 */
namespace PocketOutlook
{
    using System;

    public class TimeZone
    {
        public TimeZone(Application application,
                        ref IntPtr rpITimeZone)
        {
            m_application = application;
            m_pITimeZone = rpITimeZone;
        }

        private Application m_application;
        private IntPtr m_pITimeZone;
    }

}
