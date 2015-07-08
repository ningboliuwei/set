using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Core
{
    public class XlsCell
    {
        private object m_object;
        private bool m_isHyperLink;
        private string m_hyperLink;

        public XlsCell(object obj)
        {
            m_object = obj;
        }
        public void SetHyperLink(string url)
        {
            m_isHyperLink = true;
            m_hyperLink = url;
        }

        public object Value
        {
            get { return m_object; }
        }

        public bool IsHyperLink
        {
            get{ return m_isHyperLink;}
        }

        public string HyperLink
        {
            get { return m_hyperLink; }
        }

        public string MarkDownText
        {
            get { 
                return  IsHyperLink
                    ? string.Format("[{0}]({1})",m_object,m_hyperLink)
                    : m_object.ToString();
            }
        }

        public override string ToString()
        {
            return m_object.ToString();
        }
    }
}
