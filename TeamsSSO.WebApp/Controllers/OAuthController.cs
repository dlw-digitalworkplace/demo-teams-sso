using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;

namespace TeamsSSO.WebApp.Controllers
{
    [ApiController]
    [Authorize]
    [Route("[controller]")]
    public class OAuthController : ControllerBase
    {
        private readonly ITokenAcquisition _tokenAcquisition;

        public OAuthController(ITokenAcquisition tokenAcquisition)
        {
            _tokenAcquisition = tokenAcquisition;
        }

        public async Task<ActionResult<string>> AcquireOBOToken(string resource)
        {
            var token = await _tokenAcquisition.GetAccessTokenForUserAsync(new[] {$"{resource.TrimEnd('/')}/.default"});

            return token;
        }
    }
}
