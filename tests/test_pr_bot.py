import sys
import os
import importlib.util
import pytest

# Dynamically import pr_bot.py
spec = importlib.util.spec_from_file_location("pr_bot", os.path.join(os.path.dirname(__file__), "../scripts/pr_bot.py"))
pr_bot = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pr_bot)

def test_build_comment_basic():
    comment = pr_bot.build_comment(
        repo="test/repo",
        pr_number="1",
        head_sha="abc1234",
        ipynb_results=[("notebooks/test.ipynb", "outputs cleared ✅", ["Non-default kernel: julia"])],
        data_files=["data.csv"],
        metrics={"accuracy": 0.9}
    )
    assert "NoteGuardian" in comment
    assert "outputs cleared" in comment
    assert "data.csv" in comment
    assert "accuracy" in comment
    assert "Non-default kernel" in comment
