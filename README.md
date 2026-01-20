# Audit B - C# WinForms Application
The application connects to a Microsoft SQL Server database. Update the connection string in `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=***;Database=*****;User Id=***;Password=****;TrustServerCertificate=True;"
  },
  "AppSettings": {
    "DatabaseTimeout": 30
  }
}
```

