# PowerInsighter

A WPF application for connecting to Power BI Desktop and exploring the data model metadata.

## Features

- **Automatic Discovery**: Automatically finds running Power BI Desktop instances
- **Port Detection**: Multiple methods to detect Analysis Services ports (process scanning, port file reading, and port scanning)
- **Metadata Explorer**: View tables, columns, and relationships from your Power BI data model
- **Modern Architecture**: Built with MVVM pattern, Dependency Injection, and async/await best practices

## Requirements

- .NET 10.0
- Power BI Desktop (to connect to)

## How to Use

1. Start Power BI Desktop and open a `.pbix` file
2. Wait for the data model to fully load
3. Run PowerInsighter
4. Click "Connect to Power BI Desktop"
5. The application will automatically find and connect to your Power BI instance

## Architecture

The application follows modern .NET best practices:

### **MVVM Pattern**
- `Models`: Data classes (`ModelMetadata`)
- `ViewModels`: UI logic with INotifyPropertyChanged
- `Views`: XAML UI (MainWindow)

### **Dependency Injection**
- Services registered in `App.xaml.cs`
- Clean separation of concerns
- Testable architecture

### **Service Layer**
- `IPowerBIService`: Interface for Power BI operations
- `PowerBIService`: Implementation with multiple discovery methods

### **Features**
- Async/await throughout
- CancellationToken support
- Proper resource disposal
- Comprehensive error handling

## Project Structure

```
PowerInsighter/
??? Models/
?   ??? ModelMetadata.cs          # Data model for metadata items
??? Services/
?   ??? IPowerBIService.cs        # Service interface
?   ??? PowerBIService.cs         # Power BI discovery and connection logic
??? ViewModels/
?   ??? MainViewModel.cs          # UI logic and commands
??? Converters/
?   ??? InverseBooleanConverter.cs # XAML value converter
??? MainWindow.xaml               # Main UI
??? MainWindow.xaml.cs            # Code-behind
??? App.xaml                      # Application resources
??? App.xaml.cs                   # Application startup with DI
```

## Dependencies

- `Microsoft.AnalysisServices.NetCore.retail.amd64` (19.84.1) - Analysis Services connectivity
- `Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64` (19.84.1) - ADOMD client
- `System.Management` (10.0.4) - Process management
- `CommunityToolkit.Mvvm` (8.3.2) - MVVM helpers
- `Microsoft.Extensions.DependencyInjection` (9.0.0) - Dependency injection

## License

[Your License Here]

## Contributing

[Contributing guidelines if applicable]
