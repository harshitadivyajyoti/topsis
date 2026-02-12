# Topsis-Harshita-102317208

A Python package implementing the TOPSIS (Technique for Order Preference by Similarity to Ideal Solution) method for multi-criteria decision making.

---

##  Installation

```bash
pip install Topsis-Harshita-102317208
```

---

##  Command Line Usage

```bash
topsis <InputDataFile> <Weights> <Impacts> <OutputResultFileName>
```

### Example

```bash
topsis data.xlsx "1,1,1,2" "+,+,-,+" result.xlsx
```

---

##  Input File Requirements

- Input file must contain **three or more columns**
- First column should contain alternatives (non-numeric allowed)
- From second column onward, all values must be numeric
- Number of weights = number of impacts = number of criteria columns
- Impacts must be either '+' or '-'
- Weights and impacts must be comma separated

---

##  Output

The output file will contain:
- Topsis Score
- Rank of each alternative

---

##  Features

✔ Command line interface  
✔ Input validation  
✔ Exception handling  
✔ Excel file support  
✔ Automatic ranking  

---

##  Author

Harshita  
Roll Number: 102317208