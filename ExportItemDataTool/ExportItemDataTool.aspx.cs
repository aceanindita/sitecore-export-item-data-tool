using Newtonsoft.Json;
using Sitecore.ContentSearch;
using Sitecore.ContentSearch.Linq.Utilities;
using Sitecore.ContentSearch.SearchTypes;
using Sitecore.Data;
using Sitecore.Data.Items;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Xml;

namespace SharedSource.Verndale.ExportData
{
    public partial class ExportItemDataTool : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        { }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            var items = GetItems(txtIndexName.Text, txtLanguage.Text, txtTemplate.Text, txtLocation.Text);

            if (items != null && items.Any())
            {
                bool includeItemId = cboxListStandardFields.Items.FindByText("Item Id") != null
                                     && cboxListStandardFields.Items.FindByText("Item Id").Selected;
                bool includeItemName = cboxListStandardFields.Items.FindByText("Item Name") != null
                                       && cboxListStandardFields.Items.FindByText("Item Name").Selected;
                bool includeTemplateId = cboxListStandardFields.Items.FindByText("Template Id") != null
                                         && cboxListStandardFields.Items.FindByText("Template Id").Selected;
                bool includePath = cboxListStandardFields.Items.FindByText("Path") != null
                                   && cboxListStandardFields.Items.FindByText("Path").Selected;

                string fileContents = string.Empty;

                List<LevelField> levelFields = null;
                if (!string.IsNullOrWhiteSpace(txtFieldNames.Text))
                {
                    levelFields = GetLevelFields(txtFieldNames.Text);
                }

                #region CSV

                if (rbtnList.SelectedItem.Text == "CSV")
                {
                    const string separator = ",";
                    if (includeItemId)
                    {
                        fileContents += "Item Id" + separator;
                    }

                    if (includeItemName)
                    {
                        fileContents += "Item Name" + separator;
                    }

                    if (includeTemplateId)
                    {
                        fileContents += "Template Id" + separator;
                    }

                    if (includePath)
                    {
                        fileContents += "Path" + separator;
                    }

                    if (!string.IsNullOrWhiteSpace(txtFieldNames.Text))
                    {
                        if (levelFields == null)
                        {
                            lblStatus.Text = "Please check the format of the field string.";
                            return;
                        }

                        if (IsAdditionalLevelFieldsExists(levelFields))
                        {
                            lblStatus.Text = "CSV Format supports only 1 level of nesting due to CSV format constraints! Pick a different format!";
                            return;
                        }

                        fileContents = levelFields.Aggregate(fileContents, (current, field) => current + (field.Field + separator));
                    }

                    fileContents += "\r\n";

                    foreach (var item in items)
                    {
                        if (includeItemId)
                        {
                            fileContents += item.ID.Guid + separator;
                        }

                        if (includeItemName)
                        {
                            fileContents += item.Name + separator;
                        }

                        if (includeTemplateId)
                        {
                            fileContents += item.TemplateID.Guid + separator;
                        }

                        if (includePath)
                        {
                            fileContents += item.Paths.FullPath + separator;
                        }

                        if (!string.IsNullOrWhiteSpace(txtFieldNames.Text) && levelFields != null)
                        {
                            fileContents = levelFields.Aggregate(fileContents, (current, field) => current
                                + ((item.Fields[field.Field] != null ? "\"" + HttpUtility.HtmlDecode(item.Fields[field.Field].Value.Replace("\"", "\"\"")) + "\"" : string.Empty)
                                + separator));
                        }
                        fileContents += "\r\n";
                    }

                    Response.Clear();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", "attachment;filename=ExportedData.csv");
                    Response.Charset = "";
                    Response.ContentType = "application/text";
                    Response.Output.Write(fileContents);
                    Response.Flush();
                    Response.End();
                }

                #endregion CSV

                #region XML

                if (rbtnList.SelectedItem.Text == "XML")
                {
                    StringWriter sw = GetXml(items, includeItemId, includeItemName, includeTemplateId, includePath, levelFields);

                    Response.Clear();
                    Response.ContentType = "application/text";
                    Response.AddHeader("Content-Disposition:", "attachment;filename=ExportedData.xml");
                    Response.Output.Write(sw);
                    Response.End();

                    sw.Flush();
                    sw.Close();
                }

                #endregion XML

                #region JSON

                if (rbtnList.SelectedItem.Text == "Json")
                {
                    StringWriter sw = GetXml(items, includeItemId, includeItemName, includeTemplateId, includePath, levelFields);

                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(sw.ToString());
                    string jsonText = JsonConvert.SerializeXmlNode(doc.SelectSingleNode("/Items"));

                    sw.Flush();
                    sw.Close();

                    Response.Clear();
                    Response.ContentType = "application/text";
                    Response.AddHeader("Content-Disposition:", "attachment;filename=ExportedData.json");
                    Response.Output.Write(jsonText);
                    Response.End();
                }

                #endregion
            }
            else
            {
                lblStatus.Text = "No search results found.";
            }
        }

        private List<LevelField> GetLevelFields(string fieldStr)
        {
            List<LevelField> levelFields = new List<LevelField>();
            Stack<char> brackets = new Stack<char>();
            string currentField = "";
            string currentSubField = "";

            foreach (char c in fieldStr)
            {
                if (c == '(')
                {
                    if (brackets.Count > 0)
                    {
                        currentSubField += c;
                    }
                    brackets.Push(c);
                    continue;
                }

                if (c == ')')
                {
                    brackets.Pop();
                    if (brackets.Count > 0)
                    {
                        currentSubField += c;
                    }
                    continue;
                }

                if (c == ',' && brackets.Count == 0)
                {
                    levelFields.Add(new LevelField { Field = currentField.Trim(), SubFields = currentSubField.Trim() });
                    currentField = currentSubField = string.Empty;
                    continue;
                }

                if (brackets.Count > 0)
                {
                    currentSubField += c;
                }
                else
                {
                    currentField += c;
                }
            }
            levelFields.Add(new LevelField { Field = currentField.Trim(), SubFields = currentSubField.Trim() });

            return brackets.Count != 0 ? null : levelFields;
        }

        private bool IsAdditionalLevelFieldsExists(IEnumerable<LevelField> levelFields)
        {
            return levelFields.Any(levelField => levelField.SubFields.Contains('('));
        }

        private StringWriter GetXml(IEnumerable<Item> items, bool includeItemId, bool includeItemName,
            bool includeTemplateId, bool IncludePath, List<LevelField> levelFields)
        {
            StringWriter sw = new StringWriter();
            using (XmlWriter writer = XmlWriter.Create(sw, new XmlWriterSettings { Indent = true, Encoding = System.Text.Encoding.UTF8 }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Items"); //<Items>

                foreach (var item in items)
                {
                    writer.WriteStartElement("Item"); //<Item>
                    if (includeItemId)
                    {
                        writer.WriteStartElement("ItemId");
                        writer.WriteString(item.ID.Guid.ToString());
                        writer.WriteEndElement(); //</ItemId>
                    }

                    if (includeItemName)
                    {
                        writer.WriteStartElement("ItemName");
                        writer.WriteString(item.Name);
                        writer.WriteEndElement(); //</ItemId>
                    }

                    if (includeTemplateId)
                    {
                        writer.WriteStartElement("TemplateId");
                        writer.WriteString(item.TemplateID.Guid.ToString());
                        writer.WriteEndElement(); //</TemplateId>
                    }

                    if (IncludePath)
                    {
                        writer.WriteStartElement("Path");
                        writer.WriteString(item.Paths.FullPath);
                        writer.WriteEndElement(); //</Path>
                    }

                    if (levelFields != null)
                    {
                        foreach (var levelField in levelFields)
                        {
                            if (!string.IsNullOrWhiteSpace(levelField.Field))
                            {
                                if (!string.IsNullOrWhiteSpace(levelField.SubFields)
                                    && item.Fields[levelField.Field] != null
                                    && !string.IsNullOrWhiteSpace(item.Fields[levelField.Field].Value))
                                {
                                    foreach (var refItemId in item.Fields[levelField.Field].Value.Split('|').Where(x => !string.IsNullOrWhiteSpace(x)))
                                    {
                                        writer.WriteStartElement(levelField.Field.ToAlphanumeric());
                                        Guid tempGuid;
                                        if (Guid.TryParse(refItemId, out tempGuid))
                                        {
                                            Item tempItem = Sitecore.Context.Database.GetItem(Sitecore.Data.ID.Parse(tempGuid));
                                            if (tempItem != null)
                                            {
                                                RecursiveFieldXMLWrite(writer, tempItem, levelField.SubFields);
                                            }
                                        }
                                        writer.WriteEndElement(); //</<field>>
                                    }
                                }
                                else if (item.Fields[levelField.Field] != null)
                                {
                                    writer.WriteStartElement(levelField.Field.ToAlphanumeric());
                                    writer.WriteString(item.Fields[levelField.Field].Value);
                                    writer.WriteEndElement(); //</<field>>
                                }
                            }
                        }
                    }
                    writer.WriteEndElement(); //</Item>
                }

                writer.WriteEndElement(); //</Items>

                writer.WriteEndDocument();
                writer.Close();
            }
            return sw;
        }

        private void RecursiveFieldXMLWrite(XmlWriter writer, Item item, string fieldStr)
        {
            if (string.IsNullOrWhiteSpace(fieldStr)) return;

            List<LevelField> currentLevelFields = GetLevelFields(fieldStr);
            if (currentLevelFields != null && currentLevelFields.Any())
            {
                foreach (LevelField levelField in currentLevelFields)
                {
                    if (!string.IsNullOrWhiteSpace(levelField.Field))
                    {
                        if (!string.IsNullOrWhiteSpace(levelField.SubFields)
                            && item.Fields[levelField.Field] != null
                            && !string.IsNullOrWhiteSpace(item.Fields[levelField.Field].Value))
                        {
                            foreach (var refItemId in item.Fields[levelField.Field].Value.Split('|').Where(x => !string.IsNullOrWhiteSpace(x)))
                            {
                                writer.WriteStartElement(levelField.Field.ToAlphanumeric());
                                Guid tempGuid;
                                if (Guid.TryParse(refItemId, out tempGuid))
                                {
                                    Item tempItem = Sitecore.Context.Database.GetItem(Sitecore.Data.ID.Parse(tempGuid));
                                    if (tempItem != null)
                                    {
                                        RecursiveFieldXMLWrite(writer, tempItem, levelField.SubFields);
                                    }
                                }
                                writer.WriteEndElement(); //</<field>>
                            }
                        }
                        else if (item.Fields[levelField.Field] != null)
                        {
                            writer.WriteStartElement(levelField.Field.ToAlphanumeric());
                            writer.WriteString(item.Fields[levelField.Field].Value);
                            writer.WriteEndElement(); //</<field>>
                        }
                    }
                }
            }
        }

        public static List<Item> GetItems(string indexName, string language, string templateIds, string locationId)
        {
            using (var context = ContentSearchManager.GetIndex(indexName).CreateSearchContext())
            {
                var query = context.GetQueryable<SearchResultItem>().Where(x => x.Language == language);
                query = query.Where(x => x.Paths.Contains(Sitecore.Data.ID.Parse(locationId)));

                var templateSearch = PredicateBuilder.True<SearchResultItem>();
                if (!string.IsNullOrEmpty(templateIds))
                {
                    templateSearch = templateIds.Split(',').Where(x => !string.IsNullOrWhiteSpace(x)).Select(Sitecore.Data.ID.Parse).
                        Aggregate(templateSearch, (current, newTemplateGuid) => current.Or(t => t.TemplateId == newTemplateGuid));
                }

                query = query.Where(templateSearch);
                var count = query.Count();
                return query.Take(count).Select(x => x.GetItem()).ToList();
            }
        }
    }

    public static class Extensions
    {
        public static string ToAlphanumeric(this string param)
        {
            return (new Regex("[^a-zA-Z0-9]")).Replace(param, "");
        }
    }

    public class LevelField
    {
        public string Field { get; set; }
        public string SubFields { get; set; }
    }
}