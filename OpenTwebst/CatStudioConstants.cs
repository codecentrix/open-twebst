/*
 * This file is part of Open Twebst - web automation framework.
 * Copyright (c) 2012 Adrian Dorache
 * adrian.dorache@codecentrix.com
 *
 * Open Twebst is free software: you can redistribute it
 * and/or modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * Open Twebst is distributed in the hope that it will be
 * useful, but WITHOUT ANY WARRANTY; without even the implied warranty
 * of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with Open Twebst. If not, see <http://www.gnu.org/licenses/>.
 *
 *
 * Twebst can be used under a commercial license if such has been acquired
 * (see http://www.codecentrix.com/). The commercial license does not
 * cover derived or ported versions created by third parties under GPL.
 */

using System;
using System.Text;



namespace CatStudio
{
    class CatStudioConstants
    {
        public static readonly String OPEN_INTRO_URL              = "http://codecentrix.com/doc/aboutopentwebst.html";
        public static readonly String PAGE_STARTED_REC_URL        = "http://codecentrix.com/doc/startrec.html";
        public static readonly String TWEBST_PRODUCT_NAME         = "Open Twebst";
        public static readonly String HOOKED_BY_REC_ATTR          = "hooked_by_twebst_rec";
        public static readonly String FIND_INDEX_HELPER_ATTR      = "twebst_find_index_rec";
        public static readonly String CRNT_SELECTION_ATTR         = "twebst_crnt_selected";
        public static readonly int    MAX_TEXT_ATTR_LEN_TO_RECORD = 128;
        public static readonly String CODECENTRIX_HOME_URL        = "https://github.com/codecentrix/open-twebst";
        public static readonly String TWEBST_ONLINE_HELP_URL      = "http://codecentrix.com/doc/index.htm";
        public static readonly String TWEBST_COMMUNITY_URL        = "https://groups.google.com/forum/#!forum/twebst-automation-studio";
        public static readonly String TWEBST_LINKEDIN_URL         = "http://www.linkedin.com/groups/Twebst-3455722";
        public static readonly String TWEBST_TWITTER_URL          = "http://twitter.com/codecentrix";
        public static readonly String TWEBST_TUTORIALS_URL        = "http://codecentrix.com/doc/LnkTutorials.htm";
        public static readonly String TWEBST_NEWSLETTER_URL       = "http://eepurl.com/i-9Nr";
        public static readonly String TWEBST_EMAIL_URL            = "mailto:support@codecentrix.com";
        public static readonly String TWEBST_CHM_FILE_NAME        = "OpenTwebst.chm";
        public static readonly String TWEBST_SELECT_OUTLINE       = "1px solid black";
        public static readonly String TWEBST_SELECT_BCKG          = "LightGreen";
        public static readonly String SUPPORT_EMAIL_MESSAGE       = "Please describe the problem you have encountered ...";
        public static readonly String SUPPORT_EMAIL_SUBJECT       = "Twebst Bug Report";

        // Check for updates.
        public static readonly String TWEBST_CHECK_UPDATES_URL         = "http://codecentrix.com/doc/opentwbstupd.xml";
        public static readonly String CHECK_UPDATES_VERSION_ATTR       = "ver";
        public static readonly String CHECK_UPDATES_DOWNLOAD_URL_ATTR  = "url";

        // Intro page constants.
        public static readonly String INTRO_START_REC_ATTR      = "twbststartrec";
        public static readonly String INTRO_RUN_DEMO_ATTR       = "twbstrundemoscript";
        public static readonly String INTRO_OPEN_SAMPLES_ATTR   = "twbstopensamples";
        public static readonly String INTRO_DONT_ASK_AGAIN_ATTR = "twbstdontshowagainchckbx";


        public static bool IsWin64
        {
            get
            {
                if (!isWin64Computed)
                {
                    String programFiles32 = Environment.GetEnvironmentVariable("ProgramFiles(x86)");
                    if (programFiles32 == null)
                    {
                        isWin64 = false;
                    }
                    else
                    {
                        isWin64 = true;
                    }

                    isWin64Computed = true;
                }

                return isWin64;
            }
        }


        public static String INTRO_PAGE_URL
        {
            get
            {
                return OPEN_INTRO_URL;
            }
        }

        private static bool isWin64           = false;
        private static bool isWin64Computed   = false;
    }
}
