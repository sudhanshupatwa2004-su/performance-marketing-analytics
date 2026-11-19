"""Convert notebooks/analysis_source.py (# %% cells) into an .ipynb and execute it."""
import re, sys, nbformat
from nbclient import NotebookClient
src = open("notebooks/analysis_source.py").read()
cells = []
for block in re.split(r"^# %%", src, flags=re.M)[1:]:
    if block.startswith(" [markdown]"):
        body = block.split("\n", 1)[1]
        md = "\n".join(l[2:] if l.startswith("# ") else l.lstrip("#") for l in body.strip("\n").splitlines())
        cells.append(nbformat.v4.new_markdown_cell(md.strip()))
    else:
        cells.append(nbformat.v4.new_code_cell(block.split("\n", 1)[1].strip("\n")))
nb = nbformat.v4.new_notebook(cells=cells, metadata={"kernelspec": {"name": "python3", "display_name": "Python 3", "language": "python"}})
NotebookClient(nb, timeout=600, kernel_name="python3", resources={"metadata": {"path": "notebooks"}}).execute()
nbformat.write(nb, "notebooks/WishCare_Performance_Marketing_Analysis.ipynb")
print("ok", len(cells), "cells")
