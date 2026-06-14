from pathlib import Path
import re

def fix_file(p: Path):
    if not p.exists():
        print(f"File not found: {p}")
        return
        
    print(f"Fixing links in {p.name}...")
    content = p.read_text(encoding='utf-8', errors='ignore')
    
    # We want to replace the jiraLinks and clickUpLinks block.
    # To be line-ending and space-insensitive, we can use re.sub with a regex!
    pattern = r'jiraLinks\(displayIndex\)\s*=\s*TextValue\(\s*_\s*\n\s*ValueByHeader\(ws,\s*rowNumber,\s*"Jira Link"\)\)\s*\n\s*\n\s*clickUpLinks\(displayIndex\)\s*=\s*TextValue\(\s*_\s*\n\s*ValueByHeader\(ws,\s*rowNumber,\s*"ClickUp Link"\)\)'
    
    replacement = """If channelName = "SMS" Then
                    jiraLinks(displayIndex) = ""
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "Proof of Schedule"))
                Else
                    jiraLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "Jira Link"))
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "ClickUp Link"))
                End If"""
                
    # Normalize input content to \n
    normalized_content = content.replace('\r\n', '\n')
    
    # We use re.sub
    new_content, count = re.subn(pattern, replacement, normalized_content, flags=re.IGNORECASE)
    
    if count > 0:
        # Save back with original line endings if possible, or just \r\n
        p.write_text(new_content.replace('\n', '\r\n'), encoding='utf-8')
        print(f"  Successfully replaced {count} occurrences in {p.name}")
    else:
        print(f"  No match found in {p.name}")

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    fix_file(repo_dir / "vba_dump.txt")
    fix_file(repo_dir / "tools" / "vba_dump.txt")
    fix_file(repo_dir / "clean_temp.vba")
    fix_file(repo_dir / "tools" / "clean_mod.vba")

if __name__ == "__main__":
    main()
