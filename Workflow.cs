using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using RPA_AutoUpdateAnchor.ObjectRepository;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Excel;
using UiPath.Excel.Activities;
using UiPath.Excel.Activities.API;
using UiPath.Excel.Activities.API.Models;
using UiPath.GSuite.Activities.Api;
using UiPath.Orchestrator.Client.Models;
using UiPath.UIAutomationNext.API.Contracts;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;

namespace RPA_AutoUpdateAnchor
{
    public class Workflow : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
            // To start using services, use IntelliSense (CTRL + Space) to discover the available services:
            // e.g. system.GetAsset(...)

            // For accessing UI Elements from Object Repository, you can use the Descriptors class e.g:
            // var screen = uiAutomation.Open(Descriptors.MyApp.FirstScreen);
            // screen.Click(Descriptors.MyApp.FirstScreen.SettingsButton);
            //updateAnchor(inputKey, credentialsPath,fileId,  anchorLinks);
        }
        public string updateAnchor(string inputKey, string credentialsPath,string fileId, Dictionary<String, String> anchorLinks)
        {
            try
            {
                var credential = GoogleCredential.FromFile(credentialsPath)
                    .CreateScoped(new[] { DocsService.Scope.Documents });

                var docsService = new DocsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Docs Link Inserter"
                });

                var doc = docsService.Documents.Get(fileId).Execute();

                List<Request> requests = new List<Request>();
                HashSet<string> insertedAnchors = new HashSet<string>();
                Dictionary<string, string> remainingLinks = new Dictionary<string, string>(anchorLinks);
                string lastAnchorKey = remainingLinks.Keys.LastOrDefault();
                int totalOffset = 0;
                int paragraphIndex = 0;
                bool skipEvenCheck = false;

                foreach (var element in doc.Body.Content)
                {
                    if (element.Paragraph != null)
                    {
                        var style = element.Paragraph.ParagraphStyle;
                        if (style != null && (style.NamedStyleType == "HEADING_1" ||
                                              style.NamedStyleType == "HEADING_2" ||
                                              style.NamedStyleType == "HEADING_3" ||
                                              style.NamedStyleType == "HEADING_4"))
                            continue;

                        paragraphIndex++;

                        // Kiểm tra nếu là số lẻ hoặc đang ở chế độ bỏ qua số chẵn
                        if (!skipEvenCheck && paragraphIndex % 2 == 0)
                            continue;

                        skipEvenCheck = false; // Reset flag

                        bool foundInThisParagraph = false;
                        string text = "";
                        int startIndex = -1, endIndex = -1;

                        foreach (var textElement in element.Paragraph.Elements)
                        {
                            if (textElement.TextRun != null)
                            {
                                text += textElement.TextRun.Content;
                                if (startIndex == -1)
                                    startIndex = (textElement.StartIndex ?? 0) + totalOffset;
                                endIndex = (textElement.StartIndex ?? 0) + textElement.TextRun.Content.Length + totalOffset;
                            }
                        }

                        // Nếu không có inputKey trong đoạn số lẻ, bỏ qua kiểm tra số thứ tự và xét đoạn kế tiếp
                        if (!text.Contains(inputKey.ToLower()) && !text.Contains(inputKey.ToUpper()))
                        {
                            skipEvenCheck = true;
                            //Console.WriteLine($"Đang xử lý đoạn văn không chứa từ khóa chính {paragraphIndex}: \"{text.Trim()}\"");
                            continue;
                        }
                        else
                        {
                            //Console.WriteLine($"Đang xử lý đoạn văn chứa từ khóa chính {paragraphIndex}: \"{text.Trim()}\"");
                        }

                        foreach (var anchor in new Dictionary<string, string>(remainingLinks))
                        {
                            if (anchor.Key != lastAnchorKey && !insertedAnchors.Contains(anchor.Key) && (text.Contains(inputKey.ToLower()) || text.Contains(inputKey.ToUpper())))
                            {
                                int matchIndex = text.IndexOf(inputKey.ToLower());
                                if (matchIndex == -1)
                                {
                                    matchIndex = text.IndexOf(inputKey.ToUpper());
                                }
                                if (matchIndex != -1)
                                {
                                    int insertIndex = startIndex + matchIndex;

                                    requests.Add(new Request
                                    {
                                        DeleteContentRange = new DeleteContentRangeRequest
                                        {
                                            Range = new Google.Apis.Docs.v1.Data.Range
                                            {
                                                StartIndex = insertIndex,
                                                EndIndex = insertIndex + inputKey.Length
                                            }
                                        }
                                    });

                                    requests.Add(new Request
                                    {
                                        InsertText = new InsertTextRequest
                                        {
                                            Text = anchor.Key,
                                            Location = new Google.Apis.Docs.v1.Data.Location
                                            {
                                                Index = insertIndex
                                            }
                                        }
                                    });

                                    requests.Add(new Request
                                    {
                                        UpdateTextStyle = new UpdateTextStyleRequest
                                        {
                                            Range = new Google.Apis.Docs.v1.Data.Range
                                            {
                                                StartIndex = insertIndex,
                                                EndIndex = insertIndex + anchor.Key.Length
                                            },
                                            TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                                            {
                                                Link = new Link { Url = anchor.Value }
                                            },
                                            Fields = "link"
                                        }
                                    });

                                    //Console.WriteLine($"Thay thế {inputKey} -> {anchor.Key} và gắn link {anchor.Value}");
                                    insertedAnchors.Add(anchor.Key);
                                    remainingLinks.Remove(anchor.Key);
                                    foundInThisParagraph = true;
                                    totalOffset += anchor.Key.Length - inputKey.Length;
                                    break;
                                }
                            }
                        }

                        // Xử lý anchor cuối cùng hoặc nếu đã đến đoạn cuối mà vẫn còn anchor
                        //if (remainingLinks.Count > 0 && (lastAnchorKey == remainingLinks.Keys.First() || paragraphIndex == doc.Body.Content.Count))
                        if (!foundInThisParagraph && remainingLinks.Count > 0 && (lastAnchorKey == remainingLinks.Keys.First() || paragraphIndex == doc.Body.Content.Count))
                        {
                            foreach (var missing in new Dictionary<string, string>(remainingLinks))
                            {
                                int xemThemLength = "Xem thêm ".Length;
                                int anchorLength = missing.Key.Length + 1; // +1 vì có \n sau anchor

                                // **Chèn "Xem thêm" ngay sau đoạn văn cuối cùng**
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = "Xem thêm ",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex }
                                    }
                                });

                                // **Chèn anchor text kèm \n sau nó**
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = missing.Key + "\n",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex + xemThemLength }
                                    }
                                });

                                // **Gắn link cho anchor cuối cùng**
                                requests.Add(new Request
                                {
                                    UpdateTextStyle = new UpdateTextStyleRequest
                                    {
                                        Range = new Google.Apis.Docs.v1.Data.Range
                                        {
                                            StartIndex = endIndex + xemThemLength,
                                            EndIndex = endIndex + xemThemLength + missing.Key.Length
                                        },
                                        TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                                        {
                                            Link = new Link { Url = missing.Value }
                                        },
                                        Fields = "link"
                                    }
                                });

                                // **Định dạng đoạn văn chứa "Xem thêm <anchor>\n" thành Paragraph**
                                requests.Add(new Request
                                {
                                    UpdateParagraphStyle = new UpdateParagraphStyleRequest
                                    {
                                        Range = new Google.Apis.Docs.v1.Data.Range
                                        {
                                            StartIndex = endIndex,
                                            EndIndex = endIndex + xemThemLength + missing.Key.Length + 1
                                        },
                                        ParagraphStyle = new ParagraphStyle
                                        {
                                            NamedStyleType = "NORMAL_TEXT"
                                        },
                                        Fields = "namedStyleType"
                                    }
                                });

                                //Console.WriteLine($"Chèn anchor cuối: 'Xem thêm {missing.Key}\\n' (Link: {missing.Value})");
                                insertedAnchors.Add(missing.Key);
                                remainingLinks.Remove(missing.Key);
                                totalOffset += xemThemLength + anchorLength;
                            }
                        }

                        if (remainingLinks.Count == 0)
                            break;
                    }
                }

                if (requests.Count > 0)
                {
                    var batchUpdateRequest = new BatchUpdateDocumentRequest { Requests = requests };
                    docsService.Documents.BatchUpdate(batchUpdateRequest, fileId).Execute();
                    return "Thành công";
                }
                else
                {
                    return "Thất bại do không có request nào cập nhật anchor text";
                }
            }
            catch (Exception ex)
            {
                return "Thất bại : " + ex.Message;
                //Console.WriteLine("Error: " + ex.Message);
            }

        }
    
    
    }
}
