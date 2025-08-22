def parse_lengths(file_path):
    lengths = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("Length="):
                value = int(line.split("=")[1])
                lengths.append(value)
    return lengths


if __name__ == "__main__":
    file_path = "./FormatDealer/example.fdf" 
    result = parse_lengths(file_path)
    print(result)
