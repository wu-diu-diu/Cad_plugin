using Autodesk.AutoCAD.ApplicationServices;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace CoDesignStudy.Cad.PlugIn
{
    public class CadPlotUploader
    {
        /// <summary>
        /// 发送CAD参数和图片到指定API
        /// </summary>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="xmin">窗口最小X</param>
        /// <param name="ymin">窗口最小Y</param>
        /// <param name="xmax">窗口最大X</param>
        /// <param name="ymax">窗口最大Y</param>
        /// <param name="originx">原点X偏移</param>
        /// <param name="originy">原点Y偏移</param>
        /// <param name="serverUrl">API地址</param>
        public static async Task UploadPlotResultAsync(
            string imagePath,
            double xmin, double ymin, double xmax, double ymax,
            double originx, double originy,
            string serverUrl = "http://127.0.0.1:8000/process-cad")
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            if (!File.Exists(imagePath))
            {
                ed.WriteMessage($"\n图片文件不存在: {imagePath}");
                return;
            }

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(5); // 设置5分钟超时

                    using (var form = new MultipartFormDataContent())
                    {
                        byte[] fileBytes = File.ReadAllBytes(imagePath);
                        var fileContent = new ByteArrayContent(fileBytes);
                        fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
                        form.Add(fileContent, "files", Path.GetFileName(imagePath)); 

                        // 构造参数JSON
                        string paramJson = $@"{{
                            ""Xmin"": {xmin},
                            ""Ymin"": {ymin},
                            ""Xmax"": {xmax},
                            ""Ymax"": {ymax},
                            ""originx"": {originx},
                            ""originy"": {originy}
                        }}";

                        
                        form.Add(new StringContent(paramJson, Encoding.UTF8), "cad_params");


                        // 发送POST请求到正确的端点
                        string uploadUrl = $"{serverUrl}/upload-and-process";
                        var response = await client.PostAsync(uploadUrl, form);
                        string result = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode)
                        {
                            ed.WriteMessage($"\n✅ 图片和参数已成功发送到服务器");

                            // 解析返回的JSON结果
                            try
                            {
                                var jsonResult = JObject.Parse(result);
                                int totalImages = (int)jsonResult["total_images"];
                                int processedImages = (int)jsonResult["processed_images"];
                                bool success = (bool)jsonResult["success"];

                                ed.WriteMessage($"\n📊 处理结果: {processedImages}/{totalImages} 个图像成功");

                                //if (success && jsonResult["results"] != null)
                                //{
                                //    var results = jsonResult["results"];
                                //    foreach (var imageResult in results)
                                //    {
                                //        string imageName = imageResult.Key;
                                //        var rooms = imageResult.Value;

                                //        ed.WriteMessage($"\n📷 图片: {imageName}");
                                //        ed.WriteMessage($"🏠 发现 {rooms.Count()} 个房间:");

                                //        foreach (var room in rooms)
                                //        {
                                //            string roomName = (string)room["room_name"];
                                //            var coordinates = room["cad_coordinates"];
                                //            ed.WriteMessage($"  - {roomName}: {coordinates[0].Count()} 个坐标点");

                                //            // 这里可以在AutoCAD中创建对应的CAD对象
                                //            // CreateRoomInCAD(roomName, coordinates);
                                //        }
                                //    }
                                //}
                            }
                            catch (Exception parseEx)
                            {
                                ed.WriteMessage($"\n⚠️ 解析返回结果失败: {parseEx.Message}");
                                ed.WriteMessage($"\n原始返回: {result}");
                            }
                        }
                        else
                        {
                            ed.WriteMessage($"\n❌ 发送失败，状态码：{response.StatusCode}");
                            ed.WriteMessage($"\n错误详情：{result}");
                        }
                    }
                }
            }
            catch (HttpRequestException httpEx)
            {
                ed.WriteMessage($"\n❌ 网络请求失败: {httpEx.Message}");
                ed.WriteMessage("\n请检查服务器地址和网络连接");
            }
            catch (TaskCanceledException timeoutEx)
            {
                ed.WriteMessage($"\n❌ 请求超时: {timeoutEx.Message}");
                ed.WriteMessage("\n图片处理可能需要更长时间，请稍后重试");
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n❌ 发生异常: {ex.Message}");
                ed.WriteMessage($"\n详细错误: {ex.StackTrace}");
            }
        }
    }
}