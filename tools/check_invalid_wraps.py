from pathlib import Path
import re

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    p = repo_dir / "clean_temp.vba"
    
    content = p.read_text(encoding='utf-8', errors='ignore')
    lines = content.splitlines()
    
    invalid_keywords = ["Dim ", "For ", "Next", "If ", "Else", "Select ", "Case ", "Sub ", "Function ", "End ", "Loop"]
    
    print("Checking for invalid wraps:")
    for idx, line in enumerate(lines):
        if "On Error Resume Next:" in line:
            for kw in invalid_keywords:
                if kw in line:
                    print(f"Line {idx+1} contains '{kw}': {line.strip()}")
                    break

if __name__ == "__main__":
    main()
