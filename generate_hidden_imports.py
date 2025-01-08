import pkgutil

hidden_imports = []
for _, modname, _ in pkgutil.iter_modules():
    hidden_imports.append(modname)

print(" ".join([f"--hidden-import={mod}" for mod in hidden_imports]))
