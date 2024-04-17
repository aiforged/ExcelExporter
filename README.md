# Excel Template Populator

This project provides a tool to generate Excel files using a template. It populates sections and parameters in an Excel template with data extracted by AIForged.

## Features

- Accepts an Excel file as a template.
- Allows specification of template sections and parameters within the Excel file.
- Populates the template with data from AIForged.
- Supports Docker for containerized execution.

## How It Works

### Template Structure

The Excel template should define sections and placeholders using a specific syntax. Here's an example template structure:

Company Name: {company}

{EmployeeDetails start}
Employee ID: {emp:id}     Name: {emp:name}     Department: {emp:dept}

Contact Information:
    Email: {emp:email}
    Phone: {emp:phone}

Address:
    Street: {emp:addr:street}
    City: {emp:addr:city}
    State: {emp:addr:state}
    ZIP Code: {emp:addr:zip}

{ProjectDetails start}
Project ID: {proj:id}     Project Name: {proj:name}     Role: {proj:role}
Start Date: {proj:start}  End Date: {proj:end}
{ProjectDetails end}

{EmployeeDetails end}

### Placeholders

- **Placeholders**: Defined within curly braces `{}`. Example: `{company}`, `{emp:id}`.
- **Sections**: Marked with `start` and `end`. Example: `{EmployeeDetails start}` and `{EmployeeDetails end}`.

### Data Source

Data is populated from AIForged, which processes and extracts relevant information from documents.

## Getting Started

### Prerequisites

- .NET Core SDK
- AIForged API access

### Installation

1. Clone the repository:
```bash
   git clone https://github.com/aiforged/ExcelExporter.git
```

2. Navigate to the project directory:
```bash
   cd ExcelExporter
```

3. Restore dependencies and build the project:
```bash
   dotnet restore
   dotnet build
```

### Configuration - Development

Initialize user secrets:
`dotnet user-secrets init`

Add user secrets:
- `dotnet user-secrets set "AIForged:ApiKey" "[APIKEYHERE]"`
- `dotnet user-secrets set "AIForged:EndPoint" "[EndPoint]"`
- `dotnet user-secrets set "AIForged:ProjectId" "[ProjectId]"`
- `dotnet user-secrets set "AIForged:ServiceId" "[ServiceId]"`
- `dotnet user-secrets set "AIForged:MasterParamDefName" "[MasterParamDefName]"`
- `dotnet user-secrets set "AIForged:InputTemplatePath" "[InputTemplatePath]"`

### Running with Docker

The project includes a Dockerfile for containerized execution. You can set configuration via the `EXLGEN_CONFIG` environment variable.

Dockerfile snippet:

ENV EXLGEN_CONFIG="{'APIKey': '', 'ProjectId': , 'ServiceId': ,'AIForgedEndpoint': 'https://portal.aiforged.com','InputTemplatePath': 'Template.xlsx','OutputPath': '/exports','MasterParamDefName': ''}"

### Usage

1. Prepare your Excel template with the desired structure and placeholders.
2. Use the application to generate a populated Excel file:
   dotnet run --template "path/to/template.xlsx" --output "path/to/output.xlsx"

### EPPlus Library

This project uses the EPPlus library for handling Excel files. A commercial license is required for non-personal use. Please refer to the [EPPlus licensing](https://epplussoftware.com/developers/licenseexception) for more details.

### License

This project is licensed under the MIT License.