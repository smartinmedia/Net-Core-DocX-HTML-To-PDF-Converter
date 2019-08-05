using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    class HtmlHandler
    {
        public static string ReplaceAll(string html, Placeholders _rep)
        {
            /*
             * Replace Texts
             */
            foreach (var trDict in _rep.TextPlaceholders)
            {
                html = html.Replace(_rep.TextPlaceholderStartTag + trDict.Key + _rep.TextPlaceholderEndTag,
                    trDict.Value);
            }

            /*
             * Replace images
             */
             foreach (var replace in _rep.ImagePlaceholders)
            {
                html = html.Replace(_rep.ImagePlaceholderStartTag + replace.Key + _rep.ImagePlaceholderEndTag,
                    "<img src=\"data: image / " + ImageHandler.GetImageTypeFromMemStream(replace.Value) + "; base64," +
                    ImageHandler.GetBase64FromMemStream(replace.Value) + "\"/>");
            }

            /*
             * Replace Tables
             */
            foreach (var trDict in _rep.TablePlaceholders) //Take a Row/Table (one Dictionary) at a time
            {

                var trCol0 = trDict.First(); //This is the first placeholder in the row. We'll need 
                //just that one to find a matching table col
                // Find the first text element matching the search string - Then we will find the row -
                // where the text (placeholder) is inside a table cell --> this is the row we are searching for.
                var placeholder = _rep.TablePlaceholderStartTag + trCol0.Key + _rep.TablePlaceholderEndTag;

                var regex = new Regex("<tr((?!<tr)[\\s\\S])*"+placeholder+"[\\s\\S]*</tr>");
                var match = regex.Match(html);
                string copiedRow = match.Value;
                if (match.Success)
                {

                    for (var j = 0; j < trCol0.Value.Length; j++) //Lets create row by row and replace placeholders
                    {
                        for (var index = 0;
                            index < trDict.Count;
                            index++) //Now cycle through the "columns" (keys) of the Dictionary and replace item by item
                        {
                            var item = trDict.ElementAt(index);

                            if (html.Contains(_rep.TablePlaceholderStartTag + item.Key + _rep.TablePlaceholderEndTag))
                            {

                                html = html.Replace(
                                    _rep.TablePlaceholderStartTag + item.Key + _rep.TablePlaceholderEndTag,
                                    item.Value[index]);


                                break;
                            }
                        }





                        if (j < trCol0.Value.Length - 1)//If we have not reached the end of the rows to insert, we 
                        //can insert the resulting row
                        {
                            html = html.Insert((match.Index + (j+1)*match.Length), copiedRow);

                        }

                    }

                    html = html.Replace(_rep.NewLineTag, "<br>");

                }


            }

            return html;
        }


    }
}
