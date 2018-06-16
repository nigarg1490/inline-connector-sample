using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(TestAuthCookies.Startup))]
namespace TestAuthCookies
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
