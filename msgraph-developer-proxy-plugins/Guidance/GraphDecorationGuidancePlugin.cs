// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Extensions.Configuration;
using Microsoft.Graph.DeveloperProxy.Abstractions;
using System.CommandLine;
using System;
using Titanium.Web.Proxy.Http;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;

namespace Microsoft.Graph.DeveloperProxy.Plugins.Guidance;

public class GraphDecorationGuidancePlugin : BaseProxyPlugin {
    public override string Name => nameof(GraphDecorationGuidancePlugin);

    public override void Register(IPluginEvents pluginEvents,
                            IProxyContext context,
                            ISet<UrlToWatch> urlsToWatch,
                            IConfigurationSection? configSection = null) {
        base.Register(pluginEvents, context, urlsToWatch, configSection);

        pluginEvents.AfterResponse += AfterResponse;
    }

    private async Task AfterResponse(object? sender, ProxyResponseArgs e) {
        Request request = e.Session.HttpClient.Request;
        if (_urlsToWatch is not null && e.HasRequestUrlMatch(_urlsToWatch) && WarnNoDecoration(request))
            _logger?.LogRequest(BuildUseDecorationMessage(request), MessageType.Warning, new LoggingContext(e.Session));
    }

    private static bool WarnNoDecoration(Request request) =>
        ProxyUtils.IsGraphRequest(request) &&
        !Regex.IsMatch(request.Headers.GetFirstHeader("User-Agent").Value, "^(ISV|NONISV)\\|[A-Za-z]+\\|[A-Za-z]+\\|[0-9]+\\.[0-9]+");
        
        //
        //This pattern will match any string that starts with “ISV” or “NONISV” followed by “|” and any company name consisting of one or more alphabets followed by “|” and any application name consisting of one or more alphabets followed by “/” and any decimal number for the version.



    private static string GetDecorationParameterGuidanceUrl() => "https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online#how-to-decorate-your-http-traffic";
    private static string[] BuildUseDecorationMessage(Request r) => new[] { $"Make sure to include User Agent string in your API call to SharePoint with following naming convention", $"More info at {GetDecorationParameterGuidanceUrl()}" };
}
