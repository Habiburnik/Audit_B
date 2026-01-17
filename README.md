# Audit B - C# WinForms Application

A Windows Forms application for audit and attendance batch tracking, converted from Delphi.

## Features

- **User Authentication**: Login system for auditors
- **Batch Management**: Create and manage audit batches
- **Attendance Tracking**: Monitor employee attendance
- **Reporting**: Generate reports by auditor or batch
- **Employee Management**: Add new employees to the system

## Database Configuration

The application connects to a Microsoft SQL Server database. Update the connection string in `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=10.0.0.4,8133;Database=MISDATA;User Id=YOUR_USERNAME;Password=YOUR_PASSWORD;TrustServerCertificate=True;"
  }
}
```

Replace `YOUR_USERNAME` and `YOUR_PASSWORD` with your actual SQL Server credentials.

## Prerequisites

- .NET 8.0 SDK
- Microsoft SQL Server
- Access to MISDATA database

## Building and Running

1. Restore NuGet packages:
   ```bash
   dotnet restore
   ```

2. Build the application:
   ```bash
   dotnet build
   ```

3. Run the application:
   ```bash
   dotnet run
   ```

## Project Structure

- `Program.cs` - Application entry point
- `MainMenuForm.cs` - Main application window with login and navigation
- `AttendanceForm.cs` - Attendance management interface
- `AddNewForm.cs` - Add new employees
- `ReportByAuditorForm.cs` - Generate reports filtered by auditor
- `ReportByBatchForm.cs` - Generate reports filtered by batch
- `appsettings.json` - Database configuration

## Database Schema

The application expects the following tables in the MISDATA database:

- `Employees` - Employee information
- `Batches` - Audit batch definitions
- `Auditors` - Auditor user accounts
- `AuditSessions` - Audit session records
- `Attendance` - Employee attendance records

## Notes

This is a direct conversion from the original Delphi application, maintaining the same UI structure and functionality. All business logic has been preserved from the original queries.