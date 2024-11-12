# Getting Started with CIMS Modelling

Welcome to the CIMS Modelling repository! This guide will help you set up and run the CIMS Python package for economic climate modeling.

## Prerequisites

Before you begin, ensure you have the following installed on your computer:

- **Python 3.9+**:
  - Make sure you download and install Python version 3.9 or higher. Follow the official [installation instructions](https://wiki.python.org/moin/BeginnersGuide/Download) for your operating system.
- **Git**:
  - Git is a version control system that you will use to clone the repository. Follow the [installation instructions](https://github.com/git-guides/install-git) for your operating system.
- **Terminal or Command Prompt**
  - **Windows**: Use Command Prompt or PowerShell.
  - **macOS**: Use Terminal.
  - **Linux**: Use Terminal.

## Step-by-Step Guide

### 1. Clone the Modelling Repository

1. **Open Your Terminal**:
   - **Windows**: Open Command Prompt or PowerShell.
   - **macOS/Linux**: Open Terminal.
  
2. **Navigate to Your Desired Directory**:
   - Use the `cd` command to navigate to the directory where you want to clone the `cims-models` repository. For example:
     ```bash
     cd ~/Documents/projects/
     ```
3. **Clone the Repository**:
   - Run the following command to clone the repository:
     ```bash
     git clone https://github.com/EMRG-SFU/cims-models.git
     ```
4. **Navigate to the Cloned Directory**:
   - Change to the directory of the cloned repository:
     ```bash
     cd cims-models
     ```

### 2. Run the Launch Script

1. **Launch the Script**:
   - In the terminal, run the following command:
     ```bash
     ./launch_cims
     ```
2. **Script Actions**:
   - The script will perform the following actions:
     - **Check for Virtual Environment**:
       - If a `cims-env` virtual environment exists, it will activate it.
       - If it doesn't exist, it will create a `cims-env` virtual environment.
         - You will be prompted to select a Python version >= 3.9.
     - **Activate the Virtual Environment**:
       - The script will activate the `cims-env` virtual environment.
     - **Install or Update Dependencies**:
       - The script will install or update the `CIMS` package and all its dependencies.
     - **Launch Jupyter Lab**:
       - Finally, the script will launch Jupyter Lab.
3. **Advanced Options**:
   - For advanced users, the `launch_cims` script includes several optional parameters to customize the setup. Refer to the [Advanced Options](#5-advanced-options) section for more details.
  
### 3. Using Jupyter Lab

1. **Access Jupyter Lab**:
   - After running the script, Jupyter Lab should open in your default web browser.
   - If it doesn't open automatically, you can manually open it by navigating to `http://localhost:8888` in your web browser.

2. **Explore and Run Notebooks**:
   - By default, the `Reference.ipynb` notebook will be opened. Alternatively, you can open another notebook or create your own. 
   - Run the notebooks to begin modeling. If you're new to Jupyter, checkout [this video](https://www.youtube.com/watch?v=5pf0_bpNbkw) for an introduction. I'd suggest viewing the [Jupyter Notebook](https://www.youtube.com/watch?v=5pf0_bpNbkw&t=277s) & [Jupyter Lab](https://www.youtube.com/watch?v=5pf0_bpNbkw&t=541s) sections of the video.

3. **Exiting from Jupyter Lab**:
   - To exit Jupyter Lab, either press `Ctrl+C` in the terminal or use `File>Shut Down` in the Jupyter Lab toolbar.


### 4. Subsequent Runs

1. **Re-activate the Environment and Re-launch Jupyter Lab**:
   - If you need to re-activate the environment and re-launch Jupyter Lab, simply run the `./launch_cims` script again from the `cims-models` directory:
     ```bash
     ./launch_cims
     ```
   - This _will not_ re-clone the repository or wreck any existing setup. It will:
     - Re-activate the virtual environment.
     - Update any dependencies of the environment.
     - Launch Jupyter Lab.

## Troubleshooting

- **Permission Denied to GitHub**:
  - If you encounter a "Permission denied" error while cloning the repository, ensure you are using a Personal Access Token (PAT). Follow [this guide](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens#creating-a-personal-access-token-classic) to create and use a PAT.

- **Python Version Issues**:
  - If the launch script fails due to an incompatible Python version, ensure you have Python 3.9 or higher installed. You can download the latest version from [here](https://www.python.org/downloads/).

## Additional Resources

- **Submit Issues**: If you encounter any problems, please submit an issue on our GitHub Issues page.
- **CIMS Code Repository**: [CIMS GitHub Repository](https://github.com/EMRG-SFU/cims)

### 5. Advanced Options

The `launch_cims` script includes several optional parameters that advanced users can utilize to customize their setup. Below is a brief overview of these options:

1. **Custom Virtual Environment Name**:
   - If you want to use a custom name for the virtual environment instead of the default `cims-env`, use the `--env-name` flag:
     ```bash
     ./launch_cims --env-name my_custom_env
     ```

2. **Skip Launching Jupyter Lab**:
   - If you do not want to launch Jupyter Lab, use the `--no-jupyter` flag:
     ```bash
     ./launch_cims --no-jupyter
     ```

3. **Skip Dependency Update**:
   - To skip updating the dependencies during the launch process, use the `--skip-update` flag:
     ```bash
     ./launch_cims --skip-update
     ```

4. **Help**:
   - To see a full list of available options and their descriptions, use the `--help` flag:
     ```bash
     ./launch_cims --help
     ```
