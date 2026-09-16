using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using OutlookAddinProject.Models;

namespace OutlookAddinProject.Controllers;

public class LoginController : Controller
{
    private const string AuthenticationEndpoint =
        "https://ileana-nonsculptural-ludie.ngrok-free.dev/api/Authentication/Authenticate";

    private readonly IHttpClientFactory _httpClientFactory;

    public LoginController(IHttpClientFactory httpClientFactory)
    {
        _httpClientFactory = httpClientFactory;
    }

    [HttpGet]
    public IActionResult Index()
    {
        return View(new LoginViewModel());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Index(LoginViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            return View(model);
        }

        try
        {
            var client = _httpClientFactory.CreateClient();
            using var response = await client.PostAsJsonAsync(
                AuthenticationEndpoint,
                new { username = model.Username, password = model.Password },
                cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                HttpContext.Session.SetString("AuthenticatedUser", model.Username);
                return RedirectToAction(nameof(Success));
            }

            ModelState.AddModelError(string.Empty, await ReadErrorMessageAsync(response));
        }
        catch (HttpRequestException)
        {
            ModelState.AddModelError(string.Empty, "The authentication service is unavailable. Please try again.");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            ModelState.AddModelError(string.Empty, "The login request was cancelled. Please try again.");
        }

        return View(model);
    }

    [HttpGet]
    public IActionResult Success()
    {
        var username = HttpContext.Session.GetString("AuthenticatedUser");
        if (string.IsNullOrWhiteSpace(username))
        {
            return RedirectToAction(nameof(Index));
        }

        ViewData["Username"] = username;
        return View();
    }

    private static async Task<string> ReadErrorMessageAsync(HttpResponseMessage response)
    {
        var body = await response.Content.ReadAsStringAsync();
        if (string.IsNullOrWhiteSpace(body))
        {
            return "Invalid username or password.";
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            if (document.RootElement.TryGetProperty("message", out var message))
            {
                return message.GetString() ?? "Invalid username or password.";
            }
        }
        catch (JsonException)
        {
            // The API may return a plain-text error response.
        }

        return "Invalid username or password.";
    }
}