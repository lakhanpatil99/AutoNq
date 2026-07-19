import ast
import json
import os

def analyze_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            source = f.read()
    except Exception as e:
        return {"error": str(e)}

    tree = ast.parse(source)
    
    classes = {}
    functions = {}
    imports = []

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for name in node.names:
                imports.append(name.name)
        elif isinstance(node, ast.ImportFrom):
            imports.append(node.module)
        elif isinstance(node, ast.ClassDef):
            classes[node.name] = {
                "methods": [n.name for n in node.body if isinstance(n, ast.FunctionDef)]
            }
        elif isinstance(node, ast.FunctionDef):
            calls = []
            for subnode in ast.walk(node):
                if isinstance(subnode, ast.Call):
                    if isinstance(subnode.func, ast.Name):
                        calls.append(subnode.func.id)
                    elif isinstance(subnode.func, ast.Attribute):
                        calls.append(subnode.func.attr)
            functions[node.name] = {
                "args": [arg.arg for arg in node.args.args],
                "calls": list(set(calls)),
                "docstring": ast.get_docstring(node)
            }
            
    return {
        "imports": imports,
        "classes": classes,
        "functions": functions
    }

files_to_analyze = [
    r"d:\HACK28\app.py",
    r"d:\HACK28\excel_backend.py",
    r"d:\HACK28\mail_service.py",
    r"d:\HACK28\setup_environment.py",
    r"d:\HACK28\date_validation.py"
]

report = {}
for f in files_to_analyze:
    report[os.path.basename(f)] = analyze_file(f)

with open(r"d:\HACK28\ast_analysis.json", "w", encoding='utf-8') as out:
    json.dump(report, out, indent=2)

print("AST Analysis Complete")
