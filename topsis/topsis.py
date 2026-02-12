import sys
import pandas as pd
import numpy as np
import os

def topsis(input_file, weights, impacts, output_file):

    if not os.path.exists(input_file):
        print("Error: Input file not found")
        sys.exit(1)

    try:
        if input_file.endswith('.csv'):
            data = pd.read_csv(input_file)
        else:
            data = pd.read_excel(input_file)
    except:
        print("Error: Unable to read file")
        sys.exit(1)

    if data.shape[1] < 3:
        print("Error: Input file must contain 3 or more columns")
        sys.exit(1)

    try:
        matrix = data.iloc[:, 1:].astype(float)
    except:
        print("Error: From 2nd column onward, all values must be numeric")
        sys.exit(1)

    weights = weights.split(",")
    impacts = impacts.split(",")

    if len(weights) != len(impacts) or len(weights) != matrix.shape[1]:
        print("Error: Number of weights, impacts and criteria must be same")
        sys.exit(1)

    for i in impacts:
        if i not in ["+", "-"]:
            print("Error: Impacts must be + or -")
            sys.exit(1)

    weights = np.array(weights, dtype=float)

    norm = matrix / np.sqrt((matrix**2).sum())
    weighted = norm * weights

    ideal_best = []
    ideal_worst = []

    for i in range(len(impacts)):
        if impacts[i] == "+":
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    ideal_best = np.array(ideal_best)
    ideal_worst = np.array(ideal_worst)

    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    score = dist_worst / (dist_best + dist_worst)

    data["Topsis Score"] = score
    data["Rank"] = score.rank(ascending=False)

    data.to_excel(output_file, index=False)
    print("Result saved in", output_file)


def main():
    if len(sys.argv) != 5:
        print("Usage: python topsis.py <inputFile> <weights> <impacts> <outputFile>")
        sys.exit(1)

    topsis(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])


if __name__ == "__main__":
    main()