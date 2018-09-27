using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Dashboard_New.Startup))]
namespace Dashboard_New
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
