# Pulsar

Pulsar is a tool designed to explore the potential of AI in enhancing business processes. Developed as part of a proof of concept, this chatbot aims to provide quick, data-driven responses to basic business queries.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Technology Used

- [.Net 9](https://www.microsoft.com/net/download/windows)
- [Blazor](https://dotnet.microsoft.com/en-us/apps/aspnet/web-apps/blazor)
- [Bootstrap 5.3](https://getbootstrap.com)
- [Semantic Kernel](https://learn.microsoft.com/en-us/semantic-kernel/overview/)
- Azure OpenAI
- OpenAI

### Prerequisites

- [Visual Studio](https://www.visualstudio.com/)
- [SQL Server](https://www.microsoft.com/en-us/sql-server/sql-server-2022)
- [SQL Server Management Studio](https://msdn.microsoft.com/en-us/library/mt238290.aspx)
- [Node.js](https://nodejs.org)
- [.Net Core SDK](https://dotnet.microsoft.com/download)
- Azure Subscription (if you prefer to use Azure OpenAI)
- OpenAI (if you prefer to use OpenAI)
- DeepSeek (if you prefer to use DeepSeek)
- [WideWorldImporters DB](https://github.com/Microsoft/sql-server-samples/releases/tag/wide-world-importers-v1.0) - Sample DB

### Debugging

- Make sure the **Startup project** is `PERI.SK.Web`.
- From `PERI.SK.Web`, open `appsettings.json` and change the following accordingly...
  ```json
    {
        "ConnectionStrings": {
            "DefaultConnection": "[CONNECTION_STRING]"
        },
        "AiPlatform": {
            "Model": "[MODEL OR DEPLOYMENT NAME]",
            "Endpoint": "[ENDPOINT]",
            "ApiKey": "[KEY]",
            "Platform": "[SUPPORTED PLATFORM (see AiPlatform enum)]"
        },
        "ReferenceData": [
          {
              // The following must be handled accordingly by your business logics
              "Subject": "[SUBJECT]",
              "Reference": "[SCHEMA.TABLE (or any ID)]",
              "ConnectionStringName": "[CONNECTION STRING NAME]",
              "Type": "[TYPE OF CONNECTION (see ReferenceDataType enum)]"
          }
      ],
    }
  ```
 
### Architecture Overview

- Domain-Driven Design (DDD):
  - The architecture follows DDD principles to ensure the model reflects the business needs and provides a clean structure for complex domain logic.

- MediatR for Communication
  - **MediatR** is used as the mediator to decouple the different parts of the system. Instead of direct calls between components, all communication happens via Requests and Handlers.

- Layered Architecture:
  - The system is structured into several layers:
    - **Presentation Layer**: Exposes the chatbot’s UI.
      - `PERI.SK.Cmd` - Console App
      - `PERI.SK.Web` - Web (Blazor Server)
    - **Application Layer**: Uses **MediatR** to handle incoming commands and queries. This layer acts as the orchestrator, directing requests to the appropriate services.
    - **Domain Layer**: Encapsulates the core business logic and rules, built in line with DDD principles.    
    - **Infrastructure Layer**: Deals with external dependencies like databases, APIs, and other services. This ensures that infrastructure changes are isolated from the core logic.
      - `PERI.SK.Infrastructure` - Common services
      - `PERI.SK.Infrastructure.AzureOpenAI` - AzureOpenAI implementation
      - `PERI.SK.Infrastructure.OpenAI` - OpenAI implementation
      - `PERI.SK.Infrastructure.DeepSeek` - DeepSeek implementation
      - `PERI.SK.Infrastructure.Data` - Manages fetching data from various sources

### Deployment

**Pulsar** is currently deployed on Azure. Since this is a Proof of Concept (PoC), the deployment is intended for testing and development purposes only. There is no production environment set up yet.

To deploy **Pulsar** to Azure, follow the steps below:

- Request for Azure PublishSettings:
  - Ensure you have the latest **publishSettings** file for Azure. This file includes the necessary credentials, connection strings, and configurations for deployment. You may need to request it from the team responsible for Azure management.

- Deploy to Azure:
  - Open the publishSettings file in Visual Studio (or your preferred IDE) and deploy the application to the designated **Azure App Service**.

- Build Configuration:
  - Set the build configuration to **Release** (or **Debug** as needed) in Visual Studio before publishing.

- Set Environment Variables:
  - If there are any environment-specific variables (such as API keys, database credentials, etc.), ensure they are set within the **Azure App Service** settings.

- Post-Deployment:
  - Monitor the deployment and ensure that the chatbot is running as expected.
  - If any issues arise, refer to the **Azure Application Insights** logs or Azure DevOps build logs for troubleshooting.
 
## Misc
[![Quality gate](https://sonarcloud.io/api/project_badges/quality_gate?project=periapsys_Pulsar)](https://sonarcloud.io/summary/new_code?id=periapsys_Pulsar)
