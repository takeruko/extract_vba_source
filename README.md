# What's this project? 
Python script and Pipeline configuration file to extract VBA source files from MS Office files 
and commit/push them to a Git repository. 
It helps you manage your MS Office files with VBA macros on Azure DevOps.

# Supporting Git Hosting Services
- Microsoft Azure DevOps Service

# Usage
## Settings
### for Azure DevOps
1.	Add the [**azure-pipelines-extract-vba.yml**](/yaml/azure-pipelines-extract-vba.yml) to your Git repository.
2.  Commit and push your local repository to the Azure DevOps Server.
3.  On the Azure DevOps Server, go to the **Pipelines** and create new pipeline for your repository.
4.  At the **Configure your pipeline**, select the **Existing Azure Pipelines YAML file**, 
and then specify the **azure-pipelines-extract-vba.yml** added your Git repository as a configuration file.

### How to run
1.  When you pushed your change to your Git repository on Azure DevOps,
the Pipeline will run and then extract VBA source files into the `/vba-src` directory.