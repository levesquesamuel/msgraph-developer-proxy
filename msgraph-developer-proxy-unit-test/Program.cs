//Sample retrieved and adapted from:
//https://github.com/microsoftgraph/msgraph-training-dotnet/tree/main/app-auth/GraphAppOnlyTutorial

// <ProgramSnippet>
using System;

Console.WriteLine(".NET Graph App-only Unit test\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token (Making sure connection works)");
    Console.WriteLine("2. Make graph calls to test HTTP Header decoration");
    Console.WriteLine("3. Make graph calls to test HTTP RateLimit or Retry-After");
    Console.WriteLine("4. Make graph calls to test create overwhelming traffic");


    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // Test HTTP Header Decoration

            //Valid header
            //await TestHTTPHeaderDecorationAsync("NONISV|CompanyName|AppName/1.01");
            
            //Invalid header (Missing NONISV/ISV)
            await TestHTTPHeaderDecorationAsync("RANDOMSTRING|CompanyName|AppName/1.01");

            //Invalid header (Missing Company name)
            await TestHTTPHeaderDecorationAsync("ABC123|AppNameMissingCompanyname/1.01");

            //Invalid header (Invalid version)
            await TestHTTPHeaderDecorationAsync("NONISV|CompanyName|AppName/aaa");

            //Invalid header missing
            await TestHTTPHeaderDecorationAsync("");
            break;
        case 3:
            // Run any Graph code
            await TestHTTPRateLimitRetryAfterAsync();
            break;
        case 4:
            // Run any Graph code
            await TestCreateOverwhelmingTrafic();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}


// </ProgramSnippet>

// <InitializeGraphSnippet>
void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForAppOnlyAuth(settings);
}
// </InitializeGraphSnippet>

// <DisplayAccessTokenSnippet>
async Task DisplayAccessTokenAsync()
{
    try
    {
        var appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();
        Console.WriteLine($"App-only token: {appOnlyToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting app-only access token: {ex.Message}");
    }
}
// </DisplayAccessTokenSnippet>

async Task TestCreateOverwhelmingTrafic()
{
    await GraphHelper.TestCreateOverwhelmingTrafic();
}

async Task TestHTTPRateLimitRetryAfterAsync()
{
    await GraphHelper.TestHTTPRateLimitRetryAfterAsync();
}

async Task TestHTTPHeaderDecorationAsync(string httpHeader)
{
    await GraphHelper.TestHTTPHeaderDecorationAsync(httpHeader);
}

// <ListUsersSnippet>
async Task ListUsersAsync()
{
    try
    {
        var userPage = await GraphHelper.GetUsersAsync();

        if (userPage?.Value == null)
        {
            Console.WriteLine("No results returned.");
            return;
        }

        // Output each users's details
        foreach (var user in userPage.Value)
        {
            Console.WriteLine($"User: {user.DisplayName ?? "NO NAME"}");
            Console.WriteLine($"  ID: {user.Id}");
            Console.WriteLine($"  Email: {user.Mail ?? "NO EMAIL"}");
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // var nextPageRequest = new UsersRequestBuilder(userPage.OdataNextLink, _appClient.RequestAdapter);
        // var nextPage = await nextPageRequest.GetAsync();
        var moreAvailable = !string.IsNullOrEmpty(userPage.OdataNextLink);

        Console.WriteLine($"\nMore users available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting users: {ex.Message}");
    }
}
// </ListUsersSnippet>