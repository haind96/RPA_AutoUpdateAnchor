using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;

namespace RPA_AutoUpdateAnchor
{
    public class InsertAnchor
    {
        public static string InsertAnchorsAndOptionalImage(
        string inputKey,
        string credentialsPath,
        string fileId,
        Dictionary<string, string> anchorLinks,
        bool insertLastWithSeeMore = false,
        bool insertAllIncludingLast = false,
        string imageUrl = null)
        {
            try
            {
                var docsService = InitDocsService(credentialsPath);
                var doc = docsService.Documents.Get(fileId).Execute();
                var requests = new List<Request>();
                var inserted = new HashSet<string>();
                var remaining = new Dictionary<string, string>(anchorLinks);
                var lastAnchor = remaining.Keys.LastOrDefault();
                int offset = 0, pIndex = 0;
                bool skipEven = false;

                foreach (var el in doc.Body.Content)
                {
                    if (el.Paragraph == null || IsHeading(el.Paragraph.ParagraphStyle)) continue;
                    pIndex++;
                    if (!skipEven && pIndex % 2 == 0) continue;
                    skipEven = false;

                    string text = "";
                    int start = -1, end = -1;
                    foreach (var run in el.Paragraph.Elements)
                    {
                        if (run.TextRun != null)
                        {
                            text += run.TextRun.Content;
                            if (start == -1) start = (run.StartIndex ?? 0) + offset;
                            end = (run.StartIndex ?? 0) + run.TextRun.Content.Length + offset;
                        }
                    }

                    if (!text.ToLower().Contains(inputKey.ToLower()))
                    {
                        skipEven = true;
                        continue;
                    }

                    foreach (var a in new Dictionary<string, string>(remaining))
                    {
                        // Nếu không chèn anchor cuối, bỏ qua nó
                        if (!insertAllIncludingLast && a.Key == lastAnchor) continue;
                        if (inserted.Contains(a.Key)) continue;

                        int match = text.ToLower().IndexOf(inputKey.ToLower());
                        if (match != -1)
                        {
                            InsertAnchorRequests(requests, inputKey, a.Key, a.Value, start + match);
                            inserted.Add(a.Key);
                            remaining.Remove(a.Key);
                            offset += a.Key.Length - inputKey.Length;
                            break;
                        }
                    }

                    if (remaining.Count == 1 && remaining.ContainsKey(lastAnchor) && insertLastWithSeeMore)
                    {
                        InsertSeeMoreBlock(requests, lastAnchor, remaining[lastAnchor], end);
                        remaining.Remove(lastAnchor);
                        break;
                    }

                    if (remaining.Count == 0) break;
                }

                // Nếu vẫn còn anchor và không dùng “Xem thêm” → chèn plain cuối
                if (remaining.Count > 0 && !insertLastWithSeeMore)
                {
                    int insertIndex = -1;

                    // Tìm đoạn cuối không phải Heading
                    for (int i = doc.Body.Content.Count - 1; i >= 0; i--)
                    {
                        var content = doc.Body.Content[i];
                        if (content.Paragraph != null &&
                            !IsHeading(content.Paragraph.ParagraphStyle) &&
                            content.EndIndex.HasValue)
                        {
                            insertIndex = content.EndIndex.Value - 1; // Tránh index vượt giới hạn
                            break;
                        }
                    }

                    if (insertIndex == -1)
                        insertIndex = doc.Body.Content.LastOrDefault()?.EndIndex ?? 1;

                    foreach (var anchor in remaining)
                    {
                        string textToInsert = anchor.Key + "\n";

                        requests.Add(new Request
                        {
                            InsertText = new InsertTextRequest
                            {
                                Text = textToInsert,
                                Location = new Location { Index = insertIndex }
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
                                TextStyle = new TextStyle { Link = new Link { Url = anchor.Value } },
                                Fields = "link"
                            }
                        });

                        insertIndex += textToInsert.Length;
                    }
                }

                // Chèn hình nếu có (tổi đa 3 ảnh)
                if (!string.IsNullOrEmpty(imageUrl))
                {
                    var imageUrls = imageUrl.Split(';')
                           .Select(s => s.Trim())
                           .Where(s => !string.IsNullOrEmpty(s))
                           .Take(3)  // Lấy tối đa 3 ảnh
                           .ToList();
                    if (imageUrls.Count > 0)
                        InsertImagesAtParagraphs(requests, doc.Body.Content, imageUrls);
                }
                if (requests.Any())
                {
                    docsService.Documents.BatchUpdate(new BatchUpdateDocumentRequest { Requests = requests }, fileId).Execute();
                    return "Thành công";
                }
                return "Thất bại do không có request nào cập nhật";
            }
            catch (Exception ex)
            {
                return "Thất bại: " + ex.Message;
            }
        }

        private static DocsService InitDocsService(string credentialsPath)
        {
            var credential = GoogleCredential.FromFile(credentialsPath)
                .CreateScoped(new[] { DocsService.Scope.Documents });

            return new DocsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Docs Link Inserter"
            });
        }

        private static bool IsHeading(ParagraphStyle style) =>
            style != null && (style.NamedStyleType?.StartsWith("HEADING_") ?? false);

        private static void InsertAnchorRequests(List<Request> requests, string inputKey, string anchorKey, string anchorUrl, int insertIndex)
        {
            // Xóa inputKey
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

            // Chèn anchor
            requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = anchorKey,
                    Location = new Location { Index = insertIndex }
                }
            });

            // Gắn link
            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertIndex,
                        EndIndex = insertIndex + anchorKey.Length
                    },
                    TextStyle = new TextStyle { Link = new Link { Url = anchorUrl } },
                    Fields = "link"
                }
            });
        }

        private static void InsertSeeMoreBlock(List<Request> requests, string anchorText, string anchorUrl, int insertIndex)
        {
            const string prefix = "Xem thêm: ";
            int xemThemLength = prefix.Length;

            requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = prefix,
                    Location = new Location { Index = insertIndex }
                }
            });

            requests.Add(new Request
            {
                UpdateParagraphStyle = new UpdateParagraphStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertIndex,
                        EndIndex = insertIndex + xemThemLength + anchorText.Length + 1
                    },
                    ParagraphStyle = new ParagraphStyle { NamedStyleType = "NORMAL_TEXT" },
                    Fields = "namedStyleType"
                }
            });

            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range { StartIndex = insertIndex, EndIndex = insertIndex + 9 },
                    TextStyle = new TextStyle { Bold = true },
                    Fields = "bold"
                }
            });

            requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = anchorText + "\n",
                    Location = new Location { Index = insertIndex + xemThemLength }
                }
            });

            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertIndex + xemThemLength,
                        EndIndex = insertIndex + xemThemLength + anchorText.Length
                    },
                    TextStyle = new TextStyle { Link = new Link { Url = anchorUrl } },
                    Fields = "link"
                }
            });
        }

        private static void InsertImagesAtParagraphs(List<Request> requests, IList<StructuralElement> content, List<string> imageUrls)
        {
            int[] targetParagraphNumbers = { 3, 5, 7 };
            int imageIndex = 0;
            int paragraphCounter = 0;

            foreach (var element in content)
            {
                if (element.Paragraph != null && !IsHeading(element.Paragraph.ParagraphStyle))
                {
                    paragraphCounter++;

                    if (targetParagraphNumbers.Contains(paragraphCounter))
                    {
                        int insertIndex = element.StartIndex ?? -1;

                        if (insertIndex != -1 && imageIndex < imageUrls.Count)
                        {
                            requests.Add(new Request
                            {
                                InsertInlineImage = new InsertInlineImageRequest
                                {
                                    Location = new Location { Index = insertIndex },
                                    Uri = imageUrls[imageIndex],
                                    ObjectSize = new Size
                                    {
                                        Width = new Dimension { Magnitude = 480, Unit = "PT" },
                                        Height = new Dimension { Magnitude = 320, Unit = "PT" }
                                    }
                                }
                            });

                            imageIndex++;
                        }

                        // Dừng sau khi chèn đủ 3 ảnh (hoặc ít hơn nếu danh sách ngắn hơn)
                        if (imageIndex >= imageUrls.Count) break;
                    }
                }
            }
        }

    }
}