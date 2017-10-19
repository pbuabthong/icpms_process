# ICP-MS Custom Interpolation
Add custom interpolation for Agilent 8800 Triple Quadrupole ICP-MS

## Getting Started

This script runs only on Python3 and processes data from Excel files (*.xlsx) with a specific format from Agilent 8800-QQQ ICP-MS. Plotting a comparison graph between the instrument linear fitting and custom 1-degree interpolated spline requires Numpy and Matplotlib. However, these two packages are completely optional.

### Prerequisites

The following Python3 packages are required:

```
Openpyxl
Numpy
Scipy
Matplotlib
```

These packages can be installed manually using
```
pip3 install ...
```
or please follow the full instruction below to install all the dependencies. 

### Installing

Clone the repository

```
$ git clone https://github.com/pbuabthong/icpms_process.git
$ cd icpms_process/
```

Install dependencies

```
$ pip3 install -r requirements.txt
```

## Usage

```
python3 interp_icpms.py [-v] [-p] [-k K] [-h] element standard_filename raw_filename analyzed_filename
```

For help, use a conventional -h or --help

```
python3 interp_icpms.py -h
```

## Authors

* **Pai Pakpoom Buabthong** - *Initial work* - [pbuabthong](https://github.com/pbuabthong)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
