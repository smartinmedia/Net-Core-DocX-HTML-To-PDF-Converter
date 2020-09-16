using System.Linq;
using System.Text.RegularExpressions;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    public static class HtmlHandler
    {

        public static string ReplaceAll(string html, Placeholders _rep)
        {

            if (_rep == null)
            {
                return html;
            }

            /*
             * Replace Texts
             */
            foreach (var trDict in _rep.TextPlaceholders)
            {
                html = html.Replace(_rep.TextPlaceholderStartTag + trDict.Key + _rep.TextPlaceholderEndTag,
                    trDict.Value);
            }

            /*
             * Replace Hyperlinks
             */
            foreach (var hyperlink in _rep.HyperlinkPlaceholders)
            {
                html = html.Replace(_rep.HyperlinkPlaceholderStartTag + hyperlink.Key + _rep.HyperlinkPlaceholderEndTag,
                    $"<a href=\"{hyperlink.Value.Link}\">{hyperlink.Value.Text}</a>");
            }

            /*
             * Replace images
             */
            foreach (var replace in _rep.ImagePlaceholders)
            {
                html = html.Replace(_rep.ImagePlaceholderStartTag + replace.Key + _rep.ImagePlaceholderEndTag,
                    $"<img src=\"data:image/{replace.Value.MemStream.GetImageType()};base64,{replace.Value.MemStream.GetBase64()}\"/>");
            }

            /*
             * Replace Tables
             */
            foreach (var table in _rep.TablePlaceholders) //Take a Row/Table (one Dictionary) at a time
            {
                var tableCol0 = table.First(); //This is the first placeholder in the row, so the first columns. We'll need 
                //just that one to find a matching table col
                // Find the first text element matching the search string - Then we will find the row -
                // where the text (placeholder) is inside a table cell --> this is the row we are searching for.
                var placeholder = _rep.TablePlaceholderStartTag + tableCol0.Key + _rep.TablePlaceholderEndTag;

                var regex = new Regex("<tr((?!<tr)[\\s\\S])*" + placeholder + "[\\s\\S]*?</tr>");
                var match = regex.Match(html);
                if (match.Success) //Now we have the correct table and the row containing the placeholders
                {
                    string copiedRow = match.Value;
                    int differenceInNoCharacters = 0;

                    for (var newRow = 0; newRow < tableCol0.Value.Length; newRow++) //Lets create new row by new row and replace placeholders
                    {
                        for (var tableCol = 0;
                            tableCol < table.Count;
                            tableCol++) //Now cycle through the "columns" (keys) of the Dictionary and replace item by item
                        {
                            var colPlaceholder = table.ElementAt(tableCol);

                            if (html.Contains(_rep.TablePlaceholderStartTag + colPlaceholder.Key + _rep.TablePlaceholderEndTag))
                            {
                                var oldHtml = html;
                                html = html.Replace(
                                    _rep.TablePlaceholderStartTag + colPlaceholder.Key + _rep.TablePlaceholderEndTag,
                                    colPlaceholder.Value[newRow]);
                                differenceInNoCharacters += html.Length - oldHtml.Length;
                            }
                        }

                        if (newRow < tableCol0.Value.Length - 1)//If we have not reached the end of the rows to insert, we 
                        //can insert the resulting row
                        {
                            html = html.Insert((match.Index + ((newRow + 1) * match.Length) + differenceInNoCharacters), copiedRow);
                        }

                    }

                    html = html.Replace(_rep.NewLineTag, "<br>");
                }
            }

            return html;
        }

    }
}
