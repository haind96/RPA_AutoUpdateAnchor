using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using RPA_AutoUpdateAnchor.ObjectRepository;
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
    public class InsertAnchor
    {
        //udate all link
        public static string InsertFull(string inputKey, string credentialsPath, string fileId, Dictionary<String, String> anchorLinks, string imageUrl)
        {
            try
            {
                // Xác thực và khởi tạo dịch vụ Google Docs
                var credential = GoogleCredential.FromFile(credentialsPath)
                    .CreateScoped(new[] { DocsService.Scope.Documents });

                var docsService = new DocsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Docs Link Inserter"
                });

                // Lấy nội dung tài liệu Google Docs
                var doc = docsService.Documents.Get(fileId).Execute();

                List<Request> requests = new List<Request>(); // Danh sách các yêu cầu cập nhật tài liệu
                HashSet<string> insertedAnchors = new HashSet<string>(); // Danh sách các anchor đã được chèn
                Dictionary<string, string> remainingLinks = new Dictionary<string, string>(anchorLinks); // Danh sách anchor chưa chèn
                string lastAnchorKey = remainingLinks.Keys.LastOrDefault(); // Lấy key của anchor cuối cùng
                int totalOffset = 0; // Dùng để bù trừ vị trí do chèn/xóa văn bản
                int paragraphIndex = 0; // Đếm số thứ tự đoạn văn
                bool skipEvenCheck = false; // Biến để bỏ qua kiểm tra đoạn chẵn/lẻ
                int paragraph5StartIndex = -1; // Vị trí bắt đầu của đoạn văn thứ 5

                foreach (var element in doc.Body.Content)
                {
                    if (element.Paragraph != null)
                    {
                        // Bỏ qua các tiêu đề (Heading 1 -> Heading 4)
                        var style = element.Paragraph.ParagraphStyle;
                        if (style != null && (style.NamedStyleType == "HEADING_1" ||
                                              style.NamedStyleType == "HEADING_2" ||
                                              style.NamedStyleType == "HEADING_3" ||
                                              style.NamedStyleType == "HEADING_4"))
                            continue;

                        paragraphIndex++; // Tăng số thứ tự đoạn văn

                        if (paragraphIndex == 5)
                        {
                            // Lưu vị trí bắt đầu của đoạn văn thứ 5 để chèn hình ảnh sau này
                            paragraph5StartIndex = element.StartIndex ?? -1;
                        }

                        // Bỏ qua các đoạn chẵn (trừ khi có yêu cầu bỏ qua kiểm tra này)
                        if (!skipEvenCheck && paragraphIndex % 2 == 0)
                            continue;

                        skipEvenCheck = false; // Reset biến này mỗi khi kiểm tra đoạn mới

                        bool foundInThisParagraph = false; // Kiểm tra xem đoạn này có chứa `inputKey` không
                        string text = ""; // Nội dung đoạn văn
                        int startIndex = -1, endIndex = -1; // Vị trí bắt đầu và kết thúc của đoạn văn

                        // Lấy nội dung văn bản của đoạn văn
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

                        // Nếu đoạn văn không chứa `inputKey`, bỏ qua và tiếp tục kiểm tra đoạn kế tiếp
                        if (!text.ToLower().Contains(inputKey.ToLower()))
                        {
                            skipEvenCheck = true;
                            continue;
                        }

                        // Chèn các anchor vào đoạn văn nếu phù hợp
                        foreach (var anchor in new Dictionary<string, string>(remainingLinks))
                        {
                            if (anchor.Key != lastAnchorKey && !insertedAnchors.Contains(anchor.Key) && text.ToLower().Contains(inputKey.ToLower()))
                            {
                                int matchIndex = text.ToLower().IndexOf(inputKey.ToLower());
                                if (matchIndex != -1)
                                {
                                    int insertIndex = startIndex + matchIndex;

                                    // Xóa `inputKey`
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

                                    // Chèn anchor text mới
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

                                    // Gán link cho anchor text
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

                                    insertedAnchors.Add(anchor.Key);
                                    remainingLinks.Remove(anchor.Key);
                                    foundInThisParagraph = true;
                                    totalOffset += anchor.Key.Length - inputKey.Length;
                                    break;
                                }
                            }
                        }

                        // Chèn anchor cuối cùng nếu chưa được chèn
                        if (!foundInThisParagraph && remainingLinks.Count > 0 && (lastAnchorKey == remainingLinks.Keys.First() || paragraphIndex == doc.Body.Content.Count))
                        {
                            int lastValidIndex = doc.Body.Content.LastOrDefault()?.EndIndex ?? 0;
                            endIndex = Math.Min(endIndex, lastValidIndex);
                            if (endIndex == lastValidIndex)
                            {
                                endIndex = lastValidIndex - 50;
                            }
                            foreach (var missing in new Dictionary<string, string>(remainingLinks))
                            {
                                int xemThemLength = "Xem thêm: ".Length;
                                int anchorLength = missing.Key.Length + 1;

                                // Chèn "Xem thêm: " vào cuối đoạn văn
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = "Xem thêm: ",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex }
                                    }
                                });

                                // Định dạng văn bản về NORMAL_TEXT để tránh mất định dạng
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

                                // Bôi đậm chữ "Xem thêm"
                                requests.Add(new Request
                                {
                                    UpdateTextStyle = new UpdateTextStyleRequest
                                    {
                                        Range = new Google.Apis.Docs.v1.Data.Range
                                        {
                                            StartIndex = endIndex,
                                            EndIndex = endIndex + 9
                                        },
                                        TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                                        {
                                            Bold = true
                                        },
                                        Fields = "bold"
                                    }
                                });

                                // Chèn anchor cuối cùng kèm xuống dòng
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = missing.Key + "\n",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex + xemThemLength }
                                    }
                                });

                                // Gán link cho anchor cuối cùng
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

                                insertedAnchors.Add(missing.Key);
                                remainingLinks.Remove(missing.Key);
                                totalOffset += xemThemLength + anchorLength;
                            }
                        }

                        //Break
                        if (remainingLinks.Count == 0)
                            break;
                    }
                }

                // **Chèn hình ảnh sau khi hoàn tất chèn anchor text**
                if (!string.IsNullOrEmpty(imageUrl) && paragraph5StartIndex != -1)
                {
                    requests.Add(new Request
                    {
                        InsertInlineImage = new InsertInlineImageRequest
                        {
                            Location = new Google.Apis.Docs.v1.Data.Location
                            {
                                Index = paragraph5StartIndex
                            },
                            Uri = imageUrl,
                            ObjectSize = new Google.Apis.Docs.v1.Data.Size
                            {
                                Width = new Dimension { Magnitude = 480, Unit = "PT" },
                                Height = new Dimension { Magnitude = 320, Unit = "PT" }
                            }
                        }
                    });
                }

                // Thực thi các yêu cầu cập nhật tài liệu
                if (requests.Count > 0)
                {
                    var batchUpdateRequest = new BatchUpdateDocumentRequest { Requests = requests };
                    docsService.Documents.BatchUpdate(batchUpdateRequest, fileId).Execute();
                    return "Thành công";
                }
                else
                {
                    return "Thất bại do không có request nào cập nhật";
                }
            }
            catch (Exception ex)
            {
                return "Thất bại: " + ex.Message;
            }


        }

        //update link xem them
        public static string InsertSeeMore(string inputKey, string credentialsPath, string fileId, Dictionary<String, String> anchorLinks)
        {
            try
            {
                // Xác thực và khởi tạo dịch vụ Google Docs
                var credential = GoogleCredential.FromFile(credentialsPath)
                    .CreateScoped(new[] { DocsService.Scope.Documents });

                var docsService = new DocsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Docs Link Inserter"
                });

                // Lấy nội dung tài liệu Google Docs
                var doc = docsService.Documents.Get(fileId).Execute();

                List<Request> requests = new List<Request>(); // Danh sách các yêu cầu cập nhật tài liệu
                HashSet<string> insertedAnchors = new HashSet<string>(); // Danh sách các anchor đã được chèn
                Dictionary<string, string> remainingLinks = new Dictionary<string, string>(anchorLinks); // Danh sách anchor chưa chèn
                int totalOffset = 0; // Dùng để bù trừ vị trí do chèn/xóa văn bản
                int paragraphIndex = 0; // Đếm số thứ tự đoạn văn
                int paragraph5EndIndex = -1; // Vị trí cuối của đoạn văn thứ 5

                // **Xác định đoạn văn thứ 5**
                foreach (var element in doc.Body.Content)
                {
                    if (element.Paragraph != null)
                    {
                        // Bỏ qua các tiêu đề (Heading 1 -> Heading 4)
                        var style = element.Paragraph.ParagraphStyle;
                        if (style != null && (style.NamedStyleType == "HEADING_1" ||
                                              style.NamedStyleType == "HEADING_2" ||
                                              style.NamedStyleType == "HEADING_3" ||
                                              style.NamedStyleType == "HEADING_4"))
                            continue;

                        paragraphIndex++; // Tăng số thứ tự đoạn văn

                        if (paragraphIndex == 5)
                        {
                            // Lấy `endIndex` của đoạn văn thứ 5
                            paragraph5EndIndex = element.EndIndex ?? -1;
                            break;
                        }
                    }
                }

                // **Chèn anchor text ở đoạn 5**
                if (paragraph5EndIndex != -1 && remainingLinks.Count > 0)
                {
                    var lastAnchor = remainingLinks.Last();

                    int xemThemLength = "Xem thêm: ".Length;
                    int anchorLength = lastAnchor.Key.Length + 1;
                    int insertIndex = paragraph5EndIndex - 1;

                    // Chèn "Xem thêm: " vào cuối đoạn văn
                    requests.Add(new Request
                    {
                        InsertText = new InsertTextRequest
                        {
                            Text = "\nXem thêm: ",
                            Location = new Google.Apis.Docs.v1.Data.Location { Index = insertIndex }
                        }
                    });

                    // Định dạng văn bản về NORMAL_TEXT để tránh mất định dạng
                    requests.Add(new Request
                    {
                        UpdateParagraphStyle = new UpdateParagraphStyleRequest
                        {
                            Range = new Google.Apis.Docs.v1.Data.Range
                            {
                                StartIndex = insertIndex,
                                EndIndex = insertIndex + xemThemLength + lastAnchor.Key.Length + 1
                            },
                            ParagraphStyle = new ParagraphStyle
                            {
                                NamedStyleType = "NORMAL_TEXT"
                            },
                            Fields = "namedStyleType"
                        }
                    });

                    // Bôi đậm chữ "Xem thêm"
                    requests.Add(new Request
                    {
                        UpdateTextStyle = new UpdateTextStyleRequest
                        {
                            Range = new Google.Apis.Docs.v1.Data.Range
                            {
                                StartIndex = insertIndex,
                                EndIndex = insertIndex + 9
                            },
                            TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                            {
                                Bold = true
                            },
                            Fields = "bold"
                        }
                    });

                    // Chèn anchor cuối cùng kèm xuống dòng
                    requests.Add(new Request
                    {
                        InsertText = new InsertTextRequest
                        {
                            Text = lastAnchor.Key + "\n",
                            Location = new Google.Apis.Docs.v1.Data.Location { Index = insertIndex + xemThemLength }
                        }
                    });

                    // Gán link cho anchor cuối cùng
                    requests.Add(new Request
                    {
                        UpdateTextStyle = new UpdateTextStyleRequest
                        {
                            Range = new Google.Apis.Docs.v1.Data.Range
                            {
                                StartIndex = insertIndex + xemThemLength,
                                EndIndex = insertIndex + xemThemLength + lastAnchor.Key.Length
                            },
                            TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                            {
                                Link = new Link { Url = lastAnchor.Value }
                            },
                            Fields = "link"
                        }
                    });

                    insertedAnchors.Add(lastAnchor.Key);
                    remainingLinks.Remove(lastAnchor.Key);
                }

                // Thực thi các yêu cầu cập nhật tài liệu
                if (requests.Count > 0)
                {
                    var batchUpdateRequest = new BatchUpdateDocumentRequest { Requests = requests };
                    docsService.Documents.BatchUpdate(batchUpdateRequest, fileId).Execute();
                    return "Thành công";
                }
                else
                {
                    return "Thất bại do không có request nào cập nhật";
                }
            }
            catch (Exception ex)
            {
                return "Thất bại: " + ex.Message;
            }

        }

        //update 2 link
        public static string InsertLinkNoneSeeMore(string inputKey, string credentialsPath, string fileId, Dictionary<String, String> anchorLinks)
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

                        if (!skipEvenCheck && paragraphIndex % 2 == 0)
                            continue;

                        skipEvenCheck = false;

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

                        if (!text.ToLower().Contains(inputKey.ToLower()))
                        {
                            skipEvenCheck = true;
                            continue;
                        }

                        foreach (var anchor in new Dictionary<string, string>(remainingLinks))
                        {
                            if (!insertedAnchors.Contains(anchor.Key) && text.ToLower().Contains(inputKey.ToLower()))
                            {
                                int matchIndex = text.ToLower().IndexOf(inputKey.ToLower());

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

                                    insertedAnchors.Add(anchor.Key);
                                    remainingLinks.Remove(anchor.Key);
                                    foundInThisParagraph = true;
                                    totalOffset += anchor.Key.Length - inputKey.Length;
                                    break;
                                }
                            }
                        }

                        if (!foundInThisParagraph && remainingLinks.Count > 0 && paragraphIndex == doc.Body.Content.Count)
                        {
                            foreach (var missing in new Dictionary<string, string>(remainingLinks))
                            {
                                int xemThemLength = "Xem thêm: ".Length;
                                int anchorLength = missing.Key.Length + 1;

                                // **Chèn "Xem thêm: " ngay sau đoạn văn cuối cùng**
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = "Xem thêm: ",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex }
                                    }
                                });

                                // **Đặt NamedStyleType = "NORMAL_TEXT" trước để không làm mất định dạng**
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

                                // **Bôi đậm từ "Xem thêm" sau khi set NORMAL_TEXT**
                                requests.Add(new Request
                                {
                                    UpdateTextStyle = new UpdateTextStyleRequest
                                    {
                                        Range = new Google.Apis.Docs.v1.Data.Range
                                        {
                                            StartIndex = endIndex,
                                            EndIndex = endIndex + 9 // "Xem thêm" dài 9 ký tự
                                        },
                                        TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                                        {
                                            Bold = true
                                        },
                                        Fields = "bold"
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
            }

        }

        //update normal
        public static string InsertAnchorText(string inputKey, string credentialsPath, string fileId, Dictionary<String, String> anchorLinks)
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

                        if (!skipEvenCheck && paragraphIndex % 2 == 0)
                            continue;

                        skipEvenCheck = false;

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

                        if (!text.ToLower().Contains(inputKey.ToLower()))
                        {
                            skipEvenCheck = true;
                            continue;
                        }

                        foreach (var anchor in new Dictionary<string, string>(remainingLinks))
                        {
                            if (anchor.Key != lastAnchorKey && !insertedAnchors.Contains(anchor.Key) && text.ToLower().Contains(inputKey.ToLower()))
                            {
                                int matchIndex = text.ToLower().IndexOf(inputKey.ToLower());

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

                                    insertedAnchors.Add(anchor.Key);
                                    remainingLinks.Remove(anchor.Key);
                                    foundInThisParagraph = true;
                                    totalOffset += anchor.Key.Length - inputKey.Length;
                                    break;
                                }
                            }
                        }

                        // *** Sửa lỗi chèn "Xem thêm: " cuối tài liệu ***
                        if (!foundInThisParagraph && remainingLinks.Count > 0 &&
                            (lastAnchorKey == remainingLinks.Keys.First() || paragraphIndex == doc.Body.Content.Count))
                        {
                            // **Lấy vị trí cuối cùng hợp lệ của tài liệu**
                            int lastValidIndex = doc.Body.Content.LastOrDefault()?.EndIndex ?? 0;
                            endIndex = Math.Min(endIndex, lastValidIndex);
                            if (endIndex == lastValidIndex)
                            {
                                endIndex = lastValidIndex - 50;
                            }
                            foreach (var missing in new Dictionary<string, string>(remainingLinks))
                            {
                                int xemThemLength = "Xem thêm: ".Length;

                                // **Chèn "Xem thêm: " vào vị trí cuối cùng hợp lệ**
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = "Xem thêm: ",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex }
                                    }
                                });

                                // **Đặt NamedStyleType = "NORMAL_TEXT"**
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

                                // **Bôi đậm "Xem thêm"**
                                requests.Add(new Request
                                {
                                    UpdateTextStyle = new UpdateTextStyleRequest
                                    {
                                        Range = new Google.Apis.Docs.v1.Data.Range
                                        {
                                            StartIndex = endIndex,
                                            EndIndex = endIndex + 9 // "Xem thêm" dài 9 ký tự
                                        },
                                        TextStyle = new Google.Apis.Docs.v1.Data.TextStyle
                                        {
                                            Bold = true
                                        },
                                        Fields = "bold"
                                    }
                                });

                                // **Chèn anchor text kèm \n**
                                requests.Add(new Request
                                {
                                    InsertText = new InsertTextRequest
                                    {
                                        Text = missing.Key + "\n",
                                        Location = new Google.Apis.Docs.v1.Data.Location { Index = endIndex + xemThemLength }
                                    }
                                });

                                // **Gắn link cho anchor**
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

                                insertedAnchors.Add(missing.Key);
                                remainingLinks.Remove(missing.Key);
                                totalOffset += xemThemLength + missing.Key.Length + 1;
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
                return "Thất bại: " + ex.Message;
            }

        }

        //updateimage
        public static string InsertImage(string credentialsPath, string fileId, string imageUrl, string updateResult)
        {
            try
            {
                // Xác thực và khởi tạo dịch vụ Google Docs
                var credential = GoogleCredential.FromFile(credentialsPath)
                    .CreateScoped(new[] { DocsService.Scope.Documents });

                var docsService = new DocsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Docs Image Inserter"
                });

                // Lấy nội dung tài liệu Google Docs
                var doc = docsService.Documents.Get(fileId).Execute();

                List<Request> requests = new List<Request>(); // Danh sách các yêu cầu cập nhật tài liệu
                int paragraph5StartIndex = -1; // Vị trí bắt đầu của đoạn văn thứ 5
                int paragraphIndex = 0; // Đếm số thứ tự đoạn văn

                foreach (var element in doc.Body.Content)
                {
                    if (element.Paragraph != null)
                    {
                        // Bỏ qua các tiêu đề (Heading 1 -> Heading 4)
                        var style = element.Paragraph.ParagraphStyle;
                        if (style != null && (style.NamedStyleType == "HEADING_1" ||
                                              style.NamedStyleType == "HEADING_2" ||
                                              style.NamedStyleType == "HEADING_3" ||
                                              style.NamedStyleType == "HEADING_4"))
                            continue;

                        paragraphIndex++;

                        if (paragraphIndex == 5)
                        {
                            // Lưu vị trí bắt đầu của đoạn văn thứ 5 để chèn hình ảnh
                            paragraph5StartIndex = element.StartIndex ?? -1;
                            break; // Không cần duyệt thêm
                        }
                    }
                }

                // **Chèn hình ảnh sau khi tìm thấy đoạn văn thứ 5**
                if (!string.IsNullOrEmpty(imageUrl) && paragraph5StartIndex != -1)
                {
                    requests.Add(new Request
                    {
                        InsertInlineImage = new InsertInlineImageRequest
                        {
                            Location = new Google.Apis.Docs.v1.Data.Location
                            {
                                Index = paragraph5StartIndex
                            },
                            Uri = imageUrl,
                            ObjectSize = new Google.Apis.Docs.v1.Data.Size
                            {
                                Width = new Dimension { Magnitude = 480, Unit = "PT" },
                                Height = new Dimension { Magnitude = 320, Unit = "PT" }
                            }
                        }
                    });
                }

                // Thực thi yêu cầu cập nhật tài liệu
                if (requests.Count > 0)
                {
                    var batchUpdateRequest = new BatchUpdateDocumentRequest { Requests = requests };
                    docsService.Documents.BatchUpdate(batchUpdateRequest, fileId).Execute();
                    return "Thành công";
                }
                else
                {
                    return !string.IsNullOrEmpty(updateResult) ? updateResult + ". Không có request insert image" : "Thất bại do không có request insert image";
                }
            }
            catch (Exception ex)
            {
                return !string.IsNullOrEmpty(updateResult) ? updateResult + ". Lỗi insert image : " + ex.Message : "Thất bại: " + ex.Message;
            }

        }


    }
}