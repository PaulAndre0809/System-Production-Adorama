import re, io

def generate():
    with io.open('clean_vba_export.txt', 'r', encoding='utf-16le') as f:
        content = f.read()

    modules = re.split(r'={60}\nMODULE: (.*?)\n={60}', content)
    
    out = []
    
    # Add Systemic operational mechanics and overarching documentation
    out.append("# Systemic Operational Mechanics\n")
    out.append("## Data Flow\n- Campaigns are entered in Email Campaigns or SMS Campaigns tables.\n- Dashboard is updated via RefreshDashboard without rebuilding source tables.\n- Checkboxes and formulas manage state.\n")
    out.append("## Control Flow\n- Event modules (Worksheet_Change) trigger HandleCampaignChange which orchestrates formatting, timestamps, and Dashboard updates.\n")
    out.append("## Key Dependencies\n- win32com for python automation.\n- SharePoint for monthly calendars.\n")
    out.append("## High-Level Architecture\n- VBA standard module modEmailProductionTracker acts as the core engine. Event modules are thin delegators.\n\n")

    if len(modules) == 1:
        # No module splits found, just one big block
        process_module("modEmailProductionTracker", content, out)
    else:
        for i in range(1, len(modules), 2):
            mod_name = modules[i].strip()
            mod_code = modules[i+1]
            process_module(mod_name, mod_code, out)

    with io.open('AI Documentation Notes.md', 'w', encoding='utf-8') as f:
        f.write('\n'.join(out))

def process_module(mod_name, mod_code, out):
    # Just grab the base name before any extra info
    mod_name = mod_name.split(' ')[0]
    out.append(f"# Module / File: {mod_name}\n")
    
    # find functions/subs
    pattern = re.compile(r'^\s*(?:Public |Private )?(Sub|Function)\s+([a-zA-Z0-9_]+)\s*\((.*?)\)(?:\s*As\s+([a-zA-Z0-9_]+))?', re.MULTILINE)
    
    for match in pattern.finditer(mod_code):
        ftype, fname, fparams, fret = match.groups()
        out.append(f"## Function: {fname}")
        out.append("- **Purpose**: " + ("Entry point for " + fname if ftype == "Sub" else "Calculates or retrieves " + fname) + ".")
        out.append("- **Inputs**:")
        if fparams.strip():
            params = fparams.split(',')
            for p in params:
                p = p.strip()
                # parse param Name and Type
                p_parts = p.split(' As ')
                if len(p_parts) == 2:
                    pname = p_parts[0].replace('ByVal ', '').replace('ByRef ', '').strip()
                    ptype = p_parts[1].strip()
                    out.append(f"  - {pname} ({ptype}): Parameter {pname}")
                else:
                    out.append(f"  - {p} (Variant): Parameter {p}")
        else:
            out.append("  - None")
            
        ret_type = fret if fret else "None"
        out.append(f"- **Outputs**: {ret_type}")
        out.append("- **Dependencies**: Internal VBA components.")
        out.append(f"- **Behavior**: Executes logic for {fname}.")
        
        side_effects = "Modifies workbook state." if ftype == "Sub" else "None"
        out.append(f"- **Side Effects**: {side_effects}\n")

if __name__ == '__main__':
    generate()
