using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System;
using System.IO;
using System.Threading.Tasks;

namespace GB_NewCadPlus_IV.UploadApi.Controllers
{
    /// <summary>
    /// 路由地址 /api/graphics
    /// </summary>
    [ApiController]                              // 表示这是一个 API 控制器
    [Route("api/[controller]")]         // 路由地址 = /api/graphics
    public class GraphicsController : ControllerBase
    {
        /// <summary>
        /// 上传文件的根目录路径。 定义文件保存的根目录（你可以改成自己想要的路径，注意权限）
        /// </summary>
        /// <remarks>可更改为所需的本地路径；确保应用对该目录具有适当的读写权限并考虑跨平台兼容性。</remarks>
        private readonly string _uploadRoot = @"D:\GB_UploadApi_Files";

        /// <summary>
        /// 文件名: UploadCommand.cs
        /// </summary>
        public class UploadCommand
        {
            public int CategoryId { get; set; }
            public string CategoryType { get; set; }
            public string DisplayName { get; set; }
            public string Description { get; set; }
            public string CreatedBy { get; set; }
            public string AttributesJson { get; set; }
            public IFormFile DwgFile { get; set; }
            public IFormFile? PreviewFile { get; set; }
        }


        /// <summary>
        /// httpPost上传 dwg 文件和可选的预览图，并返回上传结果的 JSON。 处理文件上传的 POST 请求，接收 UploadCommand 对象，保存文件，并返回结果。
        /// </summary>
        /// <param name="command">上传命令对象，包含文件和相关属性信息</param>
        /// <returns>返回上传结果的 JSON 对象</returns>
        [HttpPost("upload")]             // 指定这个 Action 的 URL 后缀 upload
        public async Task<IActionResult> Upload([FromForm] UploadCommand command)
        {
            // 1. 基础校验
            if (command.DwgFile == null || command.DwgFile.Length == 0)
            {
                return BadRequest(new { success = false, message = "请提供 dwg 文件" });
            }

            // 2. 构造保存目录：根目录 / categoryType / categoryId
            var saveDir = Path.Combine(_uploadRoot, command.CategoryType, command.CategoryId.ToString());
            if (!Directory.Exists(saveDir))
            {
                Directory.CreateDirectory(saveDir);   // 自动创建目录
            }

            // 3. 生成安全的文件名：Guid + 原扩展名
            string dwgExt = Path.GetExtension(command.DwgFile.FileName);// 保留原文件扩展名
            string dwgNewName = $"{Guid.NewGuid()}{dwgExt}";// 生成新的文件名
            string dwgFullPath = Path.Combine(saveDir, dwgNewName);// 最终的完整文件路径

            // 4. 保存 dwg 文件
            using (var stream = new FileStream(dwgFullPath, FileMode.Create))
            {
                await command.DwgFile.CopyToAsync(stream);// 异步保存文件
            }

            // 5. 处理预览文件（可选）
            string previewFullPath = null;// 预览图的完整路径
            string previewNewName = null;// 预览图的新文件名
            if (command.PreviewFile != null && command.PreviewFile.Length > 0)// 如果预览文件存在且不为空
            {
                string previewExt = Path.GetExtension(command.PreviewFile.FileName);// 保留原文件扩展名
                previewNewName = $"{Guid.NewGuid()}{previewExt}";// 生成新的文件名
                previewFullPath = Path.Combine(saveDir, previewNewName);// 最终的完整文件路径
                using (var stream = new FileStream(previewFullPath, FileMode.Create))// 创建文件流并保存预览图
                {
                    await command.PreviewFile.CopyToAsync(stream);// 异步保存文件
                }
            }

            // 6. 准备返回的 JSON（现阶段先不写数据库，storageId/attrId 返回 0）
            var result = new
            {
                success = true,// 上传成功
                message = "上传成功",// 消息提示
                storageId = 0, // 储存 ID，先写死
                attrId = 0, // 属性 ID，先写死
                FilePath = dwgFullPath,// dwg 文件的完整路径
                PreviewImagePath = previewFullPath,// 预览图的完整路径（如果有）
                FileStoredName = dwgNewName,// dwg 文件的新文件名
                PreviewImageName = previewNewName// 预览图的新文件名（如果有）
            };

            return Ok(result);// 返回 200 OK 和 JSON 结果
        }
    }
}