# GitHub 源码仓库

仓库以 paperfit 目录为根，包含 Python 引擎、Skill、当前使用说明和合成测试。README 中的安装与测试命令从仓库根目录执行。

## 提交范围

提交 README、SKILL、VERSION、依赖清单、scripts、tests、当前 references 和 .github。不要提交用户论文、学校原始模板、格式配置中的原文、打印预览、输出报告、历史工作日志、虚拟环境或凭据。.gitignore 已排除这些常见路径；提交前仍应审阅待提交文件。

真实测试论文保存在本机 output 中。仓库测试使用运行时生成的合成文档，不依赖桌面文件或本机 Word/WPS。

## 自动检查

GitHub Actions 配置在 Python 3.10 和 3.12 上运行单元测试及合成流水线冒烟测试。首次推送后应查看实际运行结果。本机通过不等于远程检查已通过；这些检查也不能代替 Word/WPS 的视觉与编辑验收。

## 发布

源码推送和版本发布是两步。版本发布时可用 `python scripts/build_release.py releases` 生成对应 VERSION 的压缩包及校验文件，再作为 GitHub Release 附件上传。不要覆盖已发布版本的附件，修改内容后应增加版本号。

仓库可先设为私有。当前尚未选择项目许可证，未添加 LICENSE 文件；许可方式由项目所有者决定。

自动检查配置参考 [setup-python 官方用法](https://github.com/actions/setup-python/blob/v5/README.md) 和 [checkout 官方用法](https://github.com/actions/checkout/blob/v4/README.md)。
