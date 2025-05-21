# SnapToExcel

A Python based tool to automatically create excel sheets from Screenshot and executing SQL queries.

Works on Windows and Linux OS, Not tested on Mac OS.

## Screenshots

<details>
    <summary>Show</summary>
    
![image](./screenshots/snapToExcel.png)
![image](./screenshots/snapToExcel_result.png)
    
</details>

## Use Cases

- Taking lots of Screenshot for a project validation with client.
- Adding expected value from database to excel sheet.

## Requirements

- Install [Python 3.8+](https://www.python.org/)
- virtualenv [Virtual Environment](https://virtualenv.pypa.io/en/latest/)

## Installation

1. Create a virtual environment

    ```bash
    python3 -m virtualenv venv
    ```

2. Activate the virtual environment

    On Linux:

    ```bash
    source venv/bin/activate
    ```

    On Windows (Powershell 5+):

    ```bash
    ./venv/Scripts/activate
    ```

3. Install dependencies

    ```bash
    pip install -r requirements.txt
    ```

4. Run the application

    ```bash
    python main.py
    ```

## Contributing

- Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](LICENSE)
