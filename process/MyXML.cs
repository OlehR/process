using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

class MyXML
{
    XmlDocument doc = new XmlDocument();
    public MyXML(string varFileName = "")
    {
        if (varFileName.Trim().Length > 0)
            doc.Load(varFileName);
    }
    public string GetVar(string parKey1, string parKey2 = "")
    {
        try
        {
            if (parKey2.Length == 0)
                return doc.DocumentElement.SelectSingleNode(parKey1).InnerText.Trim();
            else
                return doc.DocumentElement.SelectSingleNode(parKey1).SelectSingleNode(parKey2).InnerText.Trim();
        }
        catch (Exception ex)
        {
            return null;
        }

    }
    public string GetAttribute(string parAttribute, string parKey1, string parKey2 = "")
    {
        try
        {
            if (parKey2.Length == 0)
                return doc.DocumentElement.SelectSingleNode(parKey1).Attributes[parAttribute].Value.Trim();
            else
                return doc.DocumentElement.SelectSingleNode(parKey1).SelectSingleNode(parKey2).Attributes[parAttribute].Value.Trim();
        }
        catch (Exception ex)
        {
            return null;
        }


    }
}
