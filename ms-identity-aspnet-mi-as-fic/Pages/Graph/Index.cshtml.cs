using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.IO;
using Microsoft.Identity.Web;
using Microsoft.Extensions.Logging;

namespace MiFicExamples.Pages.Graph
{
    [AuthorizeForScopes(Scopes = new[] { "User.Read" })]
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public IndexModel(ILogger<IndexModel> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task OnGetAsync()
        {
            var user = await _graphServiceClient.Me.GetAsync();
            ViewData["Me"] = user;
            ViewData["name"] = user?.DisplayName;

            using (var photoStream = await _graphServiceClient.Me.Photo.Content.GetAsync())
            {
                if (photoStream != null)
                {
                    MemoryStream ms = new MemoryStream();
                    photoStream.CopyTo(ms);
                    byte[] buffer = ms.ToArray();
                    ViewData["photo"] = Convert.ToBase64String(buffer);
                    
                }
                else
                {
                    ViewData["photo"] =  null;
                }
            }            
        }
    }
}

