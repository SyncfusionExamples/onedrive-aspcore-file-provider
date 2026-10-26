using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// Validate OneDrive configuration at startup so misconfiguration fails fast with a clear error.
// The provider uses app-only (client credentials) authentication, so all settings including DriveId are required.
var oneDriveSettings = builder.Configuration.GetSection("OneDriveSettings");
foreach (var key in new[] { "TenantId", "ClientId", "ClientSecret", "DriveId" })
{
    if (string.IsNullOrWhiteSpace(oneDriveSettings[key]))
    {
        throw new InvalidOperationException(
            $"OneDriveSettings:{key} is missing or empty in appsettings.json. The provider uses app-only (client credentials) authentication, so '{key}' is required. " +
            (key == "DriveId"
                ? "The /me endpoint cannot be used to resolve a drive for an application principal; the drive ID must be configured explicitly."
                : "Configure a valid Azure AD application value before starting the provider."));
    }
}

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAllOrigins", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();
app.UseCors("AllowAllOrigins");

app.MapControllers();

app.Run();
