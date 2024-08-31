<a id="readme-top"></a>
# 人事資料檢查與修正工具


<ol>
  <li><a href="#about-the-project">About The Project</a></li>
  <li><a href="#key-benefits">Key Benefits</a></li>
  <li><a href="#built-with">Built With</a></li>
  <li><a href="#getting-started">Getting Started</a></li>
  <li><a href="#usage">Usage</a></li>
  <li><a href="#contact">Contact</a></li>
</ol>

## About The Project

This project is designed to streamline the process of handling personnel data, specifically focusing on verifying and correcting information such as names, ID numbers, and birthdates from an Excel file. The main objectives include:

1. Data Validation and Correction: The system reads an Excel file containing personal information and checks the accuracy of ID numbers. It also automatically corrects any incorrect birthdate formats, ensuring data integrity.

2. Output of Corrected Data: After processing, the system generates a new Excel file with the corrected data, ready for further use or storage.

3. User-Friendly Interface: The project includes a simple and intuitive UI, making it easy for users to operate the system with minimal effort.

## Key Benefits

1. Reduced Manual Effort: Automating the validation and correction process significantly reduces the time and effort required for manual checks.

2. Increased Efficiency: The system ensures faster and more reliable processing of personnel data.

3. Error Minimization: By automating checks and corrections, the system helps prevent errors that could arise from manual data handling.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Built With

* openpyxl
* Tkinter
* CustomTkinter
  
<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Getting Started

To get a local copy up and running, follow these simple steps.

1. Create a new conda environment
   ```sh
   conda create --name Format_Checker python=3.11
   ```
   
2. Activate environment
   ```sh
   conda activate Format_Checker
   ```

3. Clone the repo
   ```sh
   git clone https://github.com/Imding1211/Format_Checker.git
   ```
   
4. Change directory
   ```sh
   cd Format_Checker
   ```
   
5. Install the required Python packages
   ```sh
   pip install -r requirements.txt
   ```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Usage

* When running it for the first time, you need to create the ChromaDB database first.
   ```sh
   python main.py populate
   ```

* After creating the database, you can start the program using the following command.
   ```sh
   python main.py run
   ```

* You can use the following example questions to test if the program is running successfully.
   ```
   What hardware setup was used for training this models?
   ```
   
* When the following message appears, it means the program is running successfully.
   ```
   1 machine with 8 NVIDIA P100 GPUs.
   ```
   
* You can place your PDF files into the "data" folder, and run the following command to populate data to the database.
   ```sh
   python main.py populate
   ```

* Or you can rebuild the database using the following command.
   ```sh
   python main.py populate --reset
  ```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Contact

Chi Heng Ting - a0986772199@gmail.com

Project Link - https://github.com/Imding1211/Format_Checker

<p align="right">(<a href="#readme-top">back to top</a>)</p>
