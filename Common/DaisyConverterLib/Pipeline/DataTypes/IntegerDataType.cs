using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class IntegerDataType : ParameterDataType
    {
        private List<string> m_ValueList;
        private List<string> m_NiceNameList;
        private int m_SelectedIndex;

        public IntegerDataType(ScriptParameter p, XmlNode node) : base(p) {
            m_Parameter = p;
            m_ValueList = new List<string>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            PopulateListFromNode(node);
        }

        public IntegerDataType(int min, int max) : base()
        {
            m_ValueList = new List<string>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            m_ValueList.Add(min.ToString());
            m_ValueList.Add(max.ToString());
            for (int i = min; i <= max; i++) {
                m_NiceNameList.Add(i.ToString());
            }
        }

        private void PopulateListFromNode(XmlNode DatatypeNode)
        {
            XmlNode EnumNode = DatatypeNode.FirstChild;
            m_ValueList.Add(EnumNode.Attributes.GetNamedItem("min").Value);
            m_ValueList.Add(EnumNode.Attributes.GetNamedItem("max").Value);
            for (int i = Convert.ToInt32(m_ValueList[0]); i <= Convert.ToInt32(m_ValueList[1]); i++)
            {
                m_NiceNameList.Add(i.ToString());
            }

        }

        public int SelectedIndex
        {
            get { return Convert.ToInt32(m_SelectedIndex); }
            set
            {
                if (m_NiceNameList.Contains(value.ToString()))
                    m_SelectedIndex = value;
                else throw new System.Exception("IndexNotInRange");
            }
        }

        //private bool SetSelectedIndexAndUpdateScript(int Index)
        //{
        //    if (Index > 0 && Index < m_ValueList.Count)
        //    {
        //        m_SelectedIndex = Index;
        //        m_Parameter.ParameterValue = SelectedItemValue;
        //        return true;
        //    }
        //    else
        //        return false;
        //}

        //public string SelectedItemValue
        //{
        //    get { return m_ValueList[m_SelectedIndex]; }
        //    set
        //    {
        //        if (value != null && m_ValueList.Contains(value))
        //            SetSelectedIndexAndUpdateScript(m_ValueList.BinarySearch(value));
        //        else throw new System.Exception("NotAbleToSelectItem");
        //    }
        //}
        public List<string> GetValues { get { return m_ValueList; } }
        //public List<string> GetNiceNames { get { return m_NiceNameList; } }

        /// <summary>
        ///  Gets and sets the index of value selected.
        /// </summary>
        //public int SelectedIndex
        //{
        //    get { return m_SelectedIndex; }
        //    set
        //    {
        //        if (value >= 0 && value < m_ValueList.Count)
        //            SetSelectedIndexAndUpdateScript(value);
        //        else throw new System.Exception("IndexNotInRange");
        //    }
        //}


        //public string SelectedItemValue
        //{
        //    get { return m_ValueList[m_SelectedIndex]; }
        //    set
        //    {
        //        if (value != null && m_ValueList.Contains(value))
        //            SetSelectedIndexAndUpdateScript(m_ValueList.BinarySearch(value));
        //        else throw new System.Exception("NotAbleToSelectItem");
        //    }
        //}


        //public string SelectedItemNiceName
        //{
        //    get { return m_NiceNameList[m_SelectedIndex]; }
        //    set
        //    {
        //        if (value != null && m_NiceNameList.Contains(value))
        //            SetSelectedIndexAndUpdateScript(m_NiceNameList.BinarySearch(value));
        //        else throw new System.Exception("NotAbleToSelectItem");
        //    }
        //}


        //private bool SetSelectedIndexAndUpdateScript(int Index)
        //{
        //    if (Index > 0 && Index < m_ValueList.Count)
        //    {
        //        m_SelectedIndex = Index;
        //        m_Parameter.ParameterValue = SelectedItemValue;
        //        return true;
        //    }
        //    else
        //        return false;
        //}

    }
}
