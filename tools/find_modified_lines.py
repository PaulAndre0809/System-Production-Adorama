from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    p = repo_dir / "clean_temp.vba"
    
    content = p.read_text(encoding='utf-8', errors='ignore')
    lines = content.splitlines()
    
    print("Modified lines in clean_temp.vba:")
    for idx, line in enumerate(lines):
        if "On Error Resume Next:" in line:
            print(f"Line {idx+1}: {line.strip()}")

if __name__ == "__main__":
    main()
