using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MVCHomework1.Startup))]
namespace MVCHomework1
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
