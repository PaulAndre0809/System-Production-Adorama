from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    p = repo_dir / "clean_temp.vba"
    
    content = p.read_text(encoding='utf-8', errors='ignore')
    lines = content.splitlines()
    
    # We want to check the part of the line after "On Error Resume Next:"
    print("Checking for real invalid wraps:")
    for idx, line in enumerate(lines):
        if "On Error Resume Next:" in line:
            # extract the actual statement that was wrapped
            parts = line.split("On Error Resume Next:")
            stmt = parts[1].split(": On Error GoTo 0")[0].strip()
            
            # check if the statement itself contains syntax keywords
            invalid_keywords = ["Dim", "For ", "Next", "If ", "Else", "Select ", "Case ", "Sub ", "Function ", "End ", "Loop"]
            for kw in invalid_keywords:
                # ignore "Next" if it's not a standalone keyword (though it shouldn't be in the statement)
                if kw in stmt:
                    print(f"Line {idx+1}: Statement '{stmt}' contains '{kw}'! Line: {line.strip()}")
                    break

if __name__ == "__main__":
    main()
