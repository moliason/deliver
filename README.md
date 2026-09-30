# deliver

通用交付打包脚本。把一个目录打成 ZIP，可检查必需文件并排除临时文件。只依赖 Python 3.9+ 标准库，不需要 Redmine、GitLab、钉钉或 Jenkins。

## 使用

```bash
python deliver.py ./release ./dist/release.zip
```

检查必需文件或目录，并排除日志：

```bash
python deliver.py ./release ./dist/release.zip \
  --required README.md --required firmware \
  --exclude '*.log' --exclude 'tmp/*'
```

`--required` 和 `--exclude` 均可重复。必需目录至少要包含一个未被排除的文件；缺少必需内容时，脚本返回非零状态。ZIP 内保留源目录下的相对路径，不包含源目录本身。默认忽略 `.git`、`.svn`、`__pycache__`、`.pyc`、`.DS_Store` 和 `Thumbs.db`。输出 ZIP 应放在源目录之外。

## 测试

```bash
python -m unittest discover -s tests -v
```
