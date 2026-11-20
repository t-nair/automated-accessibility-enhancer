import ast
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# The website image copies its Python files by name, so a new module app.py
# imports has to be added there too. Forgetting auth.py crashed every new
# container on the first deploy of sign-in; the tests passed because they run
# from the repository, where the file exists.
def copied_into_web_image():
    with open(os.path.join(ROOT, "Dockerfile.web"), encoding="utf-8") as f:
        dockerfile = f.read()

    copied = set()

    for line in re.findall(r"^COPY (.+)$", dockerfile, re.MULTILINE):
        copied.update(line.split()[:-1])

    return copied


def local_modules_imported_by(module, found):
    path = os.path.join(ROOT, module + ".py")

    if module in found or not os.path.exists(path):
        return found

    found.add(module)

    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names = [alias.name for alias in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
            names = [node.module]
        else:
            continue

        for name in names:
            local_modules_imported_by(name.split(".")[0], found)

    return found


def test_every_module_the_website_imports_is_in_its_image():
    copied = copied_into_web_image()
    needed = local_modules_imported_by("app", set())
    missing = sorted(module + ".py" for module in needed if module + ".py" not in copied)

    assert missing == [], f"Add these to the COPY line in Dockerfile.web: {missing}"
