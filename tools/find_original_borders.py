from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    p = repo_dir / "tools" / "vba_dump.txt"
    if not p.exists():
        p = repo_dir / "vba_dump.txt"
        
    content = p.read_text(encoding='utf-8', errors='ignore')
    lines = content.splitlines()
    
    print("Original borders lines:")
    for idx, line in enumerate(lines):
        if "Borders" in line:
            print(f"Line {idx+1}: {line.strip()}")

if __name__ == "__main__":
    main()
