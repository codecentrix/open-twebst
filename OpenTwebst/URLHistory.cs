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
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;



namespace CatStudio
{
    class URLHistory
    {
        public bool Load()
        {
            this.urlHistory.Clear();

            String loadFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Codecentrix");
            if (!Directory.Exists(loadFolder))
            {
                return false;
            }

            loadFolder = Path.Combine(loadFolder, "Twebst");
            if (!Directory.Exists(loadFolder))
            {
                return false;
            }

            String loadFile = Path.Combine(loadFolder, DATAFILENAME);
            if (!File.Exists(loadFile))
            {
                return false;
            }

            try
            {
                XmlDocument xmlDoc = new XmlDocument();

                xmlDoc.Load(loadFile);
                foreach (XmlNode crntNode in xmlDoc.DocumentElement.ChildNodes)
                {
                    XmlElement crntElem = crntNode as XmlElement;

                    if (crntElem != null)
                    {
                        this.urlHistory.Add(crntElem.InnerText);
                    }
                }
            }
            catch
            {
                return false;
            }

            return true;
        }


        public bool Save()
        {
            try
            {
                String saveFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Codecentrix");
                if (!Directory.Exists(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                }

                saveFolder = Path.Combine(saveFolder, "Twebst");
                if (!Directory.Exists(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                }

                String saveFile = Path.Combine(saveFolder, DATAFILENAME);
                if (File.Exists(saveFile))
                {
                    File.Delete(saveFile);
                }

                XmlDocument xmlDoc  = new XmlDocument();
                XmlElement  xmlRoot = xmlDoc.CreateElement("history");

                xmlDoc.AppendChild(xmlRoot);

                foreach (String crntUrl in this.urlHistory)
                {
                    XmlElement xmlCrntUrl = xmlDoc.CreateElement("url");
                    
                    xmlCrntUrl.InnerText = crntUrl;
                    xmlRoot.AppendChild(xmlCrntUrl);
                }

                xmlDoc.Save(saveFile);
                return true;
            }
            catch
            {
                return false;
            }
        }


        public bool AddURL(String newUrl)
        {
            if (!urlHistory.Contains(newUrl))
            {
                this.urlHistory.Add(newUrl);
                return true;
            }
            else
            {
                return false;
            }
        }


        public IList<String> Items
        {
            get { return this.urlHistory; }
        }


        private List<String> urlHistory   = new List<string>();
        private const String DATAFILENAME = "twbsturls.xml";
    }
}
