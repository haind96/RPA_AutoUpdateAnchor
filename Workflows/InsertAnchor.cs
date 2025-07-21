using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
                var service = InitDocsService(credentialsPath);
                var doc = service.Documents.Get(fileId).Execute();
                var content = doc.Body.Content;

                var requests = new List<Request>();
                var inserted = new HashSet<string>();
                var lastAnchor = anchorLinks.Keys.LastOrDefault();
                var remaining = anchorLinks
                    .Where(kv => insertAllIncludingLast || kv.Key != lastAnchor)
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                int offset = 0, paragraphIndex = 0;

                for (int i = 0; i < content.Count && remaining.Count > 0; i++)
                {
                    var el = content[i];
                    if (el.Paragraph == null || IsHeading(el.Paragraph.ParagraphStyle)) continue;

                    paragraphIndex++;
                    bool isOdd = paragraphIndex % 2 == 1;

                    string text = GetText(el);
                    int start = GetStartIndex(el) + offset;
                    int end = GetEndIndex(el) + offset;

                    bool hasKey = text.ToLower().Contains(inputKey.ToLower());

                    // Nếu đoạn lẻ không có key, xét đoạn tiếp theo
                    if (!hasKey && isOdd && i + 1 < content.Count)
                    {
                        var next = content[i + 1];
                        if (next.Paragraph != null && !IsHeading(next.Paragraph.ParagraphStyle))
                        {
                            string nextText = GetText(next);
                            if (nextText.ToLower().Contains(inputKey.ToLower()))
                            {
                                el = next;
                                text = nextText;
                                i++; paragraphIndex++;
                                start = GetStartIndex(el) + offset;
                                end = GetEndIndex(el) + offset;
                                hasKey = true;
                            }
                        }
                    }

                    if (!hasKey) continue;

                    var currentAnchor = remaining.FirstOrDefault();
                    if (string.IsNullOrEmpty(currentAnchor.Key)) break;

                    // Anchor cuối - Xem thêm
                    if (currentAnchor.Key == lastAnchor && insertLastWithSeeMore && remaining.Count == 1)
                    {
                        InsertSeeMoreFormatted(requests, end, lastAnchor, currentAnchor.Value);
                        remaining.Remove(lastAnchor);
                        break;
                    }
                    else
                    {
                        int matchIndex = text.ToLower().IndexOf(inputKey.ToLower());
                        if (matchIndex >= 0)
                        {
                            InsertAnchorRequests(requests, inputKey, currentAnchor.Key, currentAnchor.Value, start + matchIndex);
                            inserted.Add(currentAnchor.Key);
                            remaining.Remove(currentAnchor.Key);
                            offset += currentAnchor.Key.Length - inputKey.Length;
                        }
                    }
                }

                // Fallback: chèn cuối nếu còn anchor
                if (remaining.Count > 0)
                {
                    int insertIndex = GetDocumentEndIndex(content);

                    if (remaining.Count == 1 && insertLastWithSeeMore)
                    {
                        var anchor = remaining.First();
                        InsertSeeMoreFormatted(requests, insertIndex, anchor.Key, anchor.Value);
                    }
                    else
                    {
                        foreach (var anchor in remaining)
                        {
                            InsertPlainAnchor(requests, insertIndex, anchor.Key, anchor.Value);
                            insertIndex += anchor.Key.Length + 1;
                        }
                    }
                }

                if (requests.Any())
                {
                    requests = requests.Where(r => r != null).ToList();
                    service.Documents.BatchUpdate(new BatchUpdateDocumentRequest { Requests = requests }, fileId).Execute();
                    Console.WriteLine("Đã chèn anchortext thành công");
                }
                else
                {
                    Console.WriteLine("Không có anchor nào được chèn");
                }

                // Chèn hình nếu có
                if (!string.IsNullOrEmpty(imageUrl))
                {
                    //TẢI LẠI document sau khi đã update anchor
                    doc = service.Documents.Get(fileId).Execute();
                    content = doc.Body.Content;
                    var imageUrls = imageUrl.Split(';').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).Take(3).ToList();
                    if (imageUrls.Count > 0)
                    {
                        var imageRequests = new List<Request>();
                        InsertImagesAfterHeading2(imageRequests, content, imageUrls);
                        if (imageRequests.Any())
                        {
                            service.Documents.BatchUpdate(new BatchUpdateDocumentRequest { Requests = imageRequests }, fileId).Execute();
                            Console.WriteLine("Đã chèn ảnh thành công.");
                        }
                        else
                        {
                            Console.WriteLine("Không có ảnh nào được chèn.");
                        }
                    }
                }

                return "Thành công";

            }
            catch (Exception ex)
            {
                return "Thất bại: " + ex.Message;
            }
        }

        private static string GetText(StructuralElement el) =>
            string.Join("", el.Paragraph.Elements.Where(e => e.TextRun != null).Select(e => e.TextRun.Content));

        private static int GetStartIndex(StructuralElement el) =>
            el.Paragraph.Elements.FirstOrDefault(e => e.TextRun != null)?.StartIndex ?? 0;

        private static int GetEndIndex(StructuralElement el) =>
            el.Paragraph.Elements.LastOrDefault(e => e.TextRun != null) is var last && last != null
                ? (last.StartIndex ?? 0) + last.TextRun.Content.Length
                : 0;

        private static int GetDocumentEndIndex(IList<StructuralElement> content)
        {
            for (int i = content.Count - 1; i >= 0; i--)
            {
                var el = content[i];
                if (el.Paragraph != null && !IsHeading(el.Paragraph.ParagraphStyle) && el.EndIndex.HasValue)
                    return el.EndIndex.Value - 1;
            }
            return content.LastOrDefault()?.EndIndex ?? 1;
        }

        private static void InsertPlainAnchor(List<Request> requests, int index, string text, string url)
        {
            requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = text + "\n",
                    Location = new Location { Index = index }
                }
            });

            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = index,
                        EndIndex = index + text.Length
                    },
                    TextStyle = new TextStyle { Link = new Link { Url = url } },
                    Fields = "link"
                }
            });
        }

        private static void InsertSeeMoreFormatted(List<Request> requests, int insertAt, string anchorText, string anchorUrl)
        {
            string boldText = "Xem thêm: ";
            string fullText = boldText + anchorText + "\n";

            // 1. Chèn văn bản đầy đủ vào cuối đoạn
            requests.Add(new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = fullText,
                    Location = new Location { Index = insertAt }
                }
            });

            // 2. Ép đoạn văn này về NORMAL_TEXT
            requests.Add(new Request
            {
                UpdateParagraphStyle = new UpdateParagraphStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertAt,
                        EndIndex = insertAt + fullText.Length
                    },
                    ParagraphStyle = new ParagraphStyle
                    {
                        NamedStyleType = "NORMAL_TEXT"
                    },
                    Fields = "namedStyleType"
                }
            });

            // 3. Bôi đậm "Xem thêm: "
            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertAt,
                        EndIndex = insertAt + boldText.Length
                    },
                    TextStyle = new TextStyle { Bold = true },
                    Fields = "bold"
                }
            });

            // 4. Gắn link cho phần anchor
            requests.Add(new Request
            {
                UpdateTextStyle = new UpdateTextStyleRequest
                {
                    Range = new Google.Apis.Docs.v1.Data.Range
                    {
                        StartIndex = insertAt + boldText.Length,
                        EndIndex = insertAt + boldText.Length + anchorText.Length
                    },
                    TextStyle = new TextStyle { Link = new Link { Url = anchorUrl } },
                    Fields = "link"
                }
            });
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

        private static void InsertImagesAfterHeading2(List<Request> requests, IList<StructuralElement> content, List<string> imageUrls)
        {
            int imageIndex = 0;
            var usedIndexes = new HashSet<int>();

            // Ưu tiên chèn sau đoạn văn đầu tiên sau HEADING_2
            for (int i = 0; i < content.Count && imageIndex < imageUrls.Count; i++)
            {
                var element = content[i];
                var paragraphStyle = element.Paragraph?.ParagraphStyle;

                if (paragraphStyle != null && paragraphStyle.NamedStyleType == "HEADING_2")
                {
                    int? insertIndex = GetFirstNonHeadingParagraphIndex(content, i);
                    if (insertIndex.HasValue && usedIndexes.Add(insertIndex.Value))
                    {
                        int adjustedIndex = insertIndex.Value;
                        if (imageIndex == imageUrls.Count - 1) adjustedIndex += 1;
                        AddImageRequest(requests, adjustedIndex, imageUrls[imageIndex]);
                        //Console.WriteLine($"Chèn ảnh sau đoạn sau HEADING_2: {imageUrls[imageIndex]}");
                        imageIndex++;
                    }
                }
            }

            // Nếu còn ảnh thì chèn sau đoạn văn đầu tiên sau HEADING_3
            for (int i = 0; i < content.Count && imageIndex < imageUrls.Count; i++)
            {
                var element = content[i];
                var paragraphStyle = element.Paragraph?.ParagraphStyle;

                if (paragraphStyle != null && paragraphStyle.NamedStyleType == "HEADING_3")
                {
                    int? insertIndex = GetFirstNonHeadingParagraphIndex(content, i);
                    if (insertIndex.HasValue && usedIndexes.Add(insertIndex.Value))
                    {
                        int adjustedIndex = insertIndex.Value;
                        if (imageIndex == imageUrls.Count - 1) adjustedIndex += 1;
                        AddImageRequest(requests, adjustedIndex, imageUrls[imageIndex]);
                        Console.WriteLine($"Chèn ảnh sau đoạn sau HEADING_3: {imageUrls[imageIndex]}");
                        imageIndex++;
                    }
                }
            }

            if (imageIndex < imageUrls.Count)
            {
                Console.WriteLine($"Còn {imageUrls.Count - imageIndex} ảnh chưa chèn do thiếu đoạn văn sau Heading 2/3.");
            }
        }

        // Trả về chỉ số đoạn văn đầu tiên sau heading, loại trừ heading
        private static int? GetFirstNonHeadingParagraphIndex(IList<StructuralElement> content, int currentIndex)
        {
            for (int j = currentIndex + 1; j < content.Count; j++)
            {
                var para = content[j].Paragraph;
                if (para != null && !IsHeading(para.ParagraphStyle))
                {
                    return content[j].EndIndex; //dùng EndIndex để chèn sau đoạn
                }
            }
            return null;
        }

        private static void CleanTrailingSpecialChars(IList<StructuralElement> content, int insertIndex)
        {
            var element = content.FirstOrDefault(e => e.StartIndex == insertIndex);
            var textRun = element?.Paragraph?.Elements?.FirstOrDefault()?.TextRun;

            if (textRun != null && !string.IsNullOrEmpty(textRun.Content))
            {
                string cleaned = Regex.Replace(textRun.Content, @"([^\w\s]+)(\s*)$", "");
                textRun.Content = cleaned + "\n"; // thêm newline lại nếu cần
            }
        }
        
        private static void AddImageRequest(List<Request> requests, int insertIndex, string imageUrl)
        {
            // 1. Thêm newline trước khi chèn ảnh để đảm bảo ảnh nằm trên dòng mới
            //            requests.Add(new Request
            //            {
            //                InsertText = new InsertTextRequest
            //                {
            //                    Text = "\n",
            //                    Location = new Location { Index = insertIndex }
            //                }
            //            });

            // 2. Chèn ảnh ngay sau newline
            requests.Add(new Request
            {
                InsertInlineImage = new InsertInlineImageRequest
                {
                    //Location = new Location { Index = insertIndex + 1 },
                    Location = new Location { Index = insertIndex },
                    Uri = imageUrl,
                    ObjectSize = new Size
                    {
                        Width = new Dimension { Magnitude = 480, Unit = "PT" },
                        Height = new Dimension { Magnitude = 320, Unit = "PT" }
                    }
                }
            });

            // 3. (Tùy chọn) thêm dòng trống sau ảnh để tránh ảnh dính sát phần dưới
            //            requests.Add(new Request
            //            {
            //                InsertText = new InsertTextRequest
            //                {
            //                    Text = "\n",
            //                    Location = new Location { Index = insertIndex + 2 }
            //                }
            //            });
        }


    }
}