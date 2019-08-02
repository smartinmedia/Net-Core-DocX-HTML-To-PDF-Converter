using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DocXToPdfConverter.DocXToPdfHandlers
{
    class HtmlHandler
    {
        public static string ReplaceAll(string html, ReplacementDictionaries _rep)
        {

            foreach (var trDict in _rep.TextReplacements)
            {
                html = html.Replace(_rep.TextReplacementStartTag + trDict.Key + _rep.TextReplacementEndTag,
                    trDict.Value);
            }

            foreach (var replace in _rep.ImageReplacements)
            {
                html = html.Replace(_rep.ImageReplacementStartTag + replace.Key + _rep.ImageReplacementEndTag,
                    "<img src=\"data: image / " + ImageHandler.GetImageTypeFromMemStream(replace.Value) + "; base64," +
                    ImageHandler.GetBase64FromMemStream(replace.Value) + "\"/>");
            }

            foreach (var trDict in _rep.TableReplacements) //Take a Row/Table (one Dictionary) at a time
            {

                var trCol0 = trDict.First(); //This is the first placeholder
                // Find the first text element matching the search string - Then we will find the row -
                // where the text (placeholder) is inside a table cell --> this is the row we are searching for.
                var placeholder = _rep.TableReplacementStartTag + trCol0.Key + _rep.TableReplacementEndTag;
                var regex = new Regex(@"<tr.*?" + placeholder + ".*?</tr>");
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

                            if (html.Contains(_rep.TableReplacementStartTag + item.Key + _rep.TableReplacementEndTag))
                            {

                                html = html.Replace(
                                    _rep.TableReplacementStartTag + item.Key + _rep.TableReplacementEndTag,
                                    item.Value[j]);


                                break;
                            }
                        }





                        if (j < trCol0.Value.Length - 1)
                        {
                            html = html.Insert((match.Index + match.Length), copiedRow);

                        }

                    }

                    html = html.Replace(_rep.NewLineTag, "<br>");

                }


            }

            return html;
        }


    }
}
