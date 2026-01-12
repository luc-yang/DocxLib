# DocxLib 贡献指南

感谢您对 DocxLib 项目的关注！

## 如何贡献

### 报告 Bug

如果您发现了 Bug，请：

1. 在 GitHub Issues 中搜索是否已有相关问题
2. 如果没有，创建一个新的 Issue，包含：
   - 清晰的标题
   - 详细的复现步骤
   - 期望的行为
   - 实际的行为
   - 环境信息（Python 版本、操作系统等）
   - 相关代码或错误日志

### 提出新功能

如果您有新功能的建议：

1. 先在 GitHub Issues 中讨论您的想法
2. 说明使用场景和预期效果
3. 等待维护者反馈

### 提交代码

#### 准备工作

1. Fork 项目仓库
2. 克隆您的 Fork：
   ```bash
   git clone https://github.com/luc-yang/DocxLib.git
   cd docxlib
   ```
3. 创建开发分支：
   ```bash
   git checkout -b feature/your-feature-name
   ```

#### 代码规范

- 遵循 PEP 8 代码规范
- 使用有意义的变量和函数名
- 添加必要的注释和文档字符串
- 保持函数简洁，单一职责

#### 提交规范

提交信息格式：
```
<type>: <description>

[可选的详细描述]
```

类型（type）：
- `feat`: 新功能
- `fix`: Bug 修复
- `docs`: 文档更新
- `style`: 代码格式调整
- `refactor`: 代码重构
- `test`: 测试相关
- `chore`: 构建/工具相关

示例：
```
feat: 添加表格合并功能

支持将多个表格合并为一个，并保持格式一致。
```

#### 测试

- 为新功能添加单元测试
- 确保所有测试通过：
  ```bash
  pytest tests/ -v
  ```
- 运行基础功能测试：
  ```bash
  python tests/test_basic.py
  ```

#### 提交 Pull Request

1. 推送到您的 Fork：
   ```bash
   git push origin feature/your-feature-name
   ```
2. 在 GitHub 上创建 Pull Request
3. 填写 PR 模板：
   - 描述改动内容
   - 关联相关 Issue
   - 确认测试通过
   - 添加截图（如适用）

### 代码审查

- 保持开放心态，接受反馈
- 及时回应审查意见
- 解释设计决策
- 根据反馈调整代码

## 开发环境设置

### 安装开发依赖

```bash
pip install -e ".[dev]"
```

### 运行测试

```bash
# 运行所有测试
pytest tests/ -v

# 运行特定测试
pytest tests/test_document.py -v

# 生成覆盖率报告
pytest tests/ --cov=docxlib --cov-report=html
```

### 代码格式化

使用 black 格式化代码：

```bash
pip install black
black docxlib/
```

### 代码检查

使用 flake8 检查代码：

```bash
pip install flake8
flake8 docxlib/
```

## 项目结构

```
docxlib/
├── docxlib/          # 主要代码
├── examples/         # 示例代码
├── tests/            # 测试代码
├── fixtures/         # 测试数据
└── docs/             # 文档（未来）
```

## 版本发布

版本号遵循语义化版本（Semantic Versioning）：

- `MAJOR.MINOR.PATCH`
- MAJOR: 不兼容的 API 变更
- MINOR: 向后兼容的新功能
- PATCH: 向后兼容的 Bug 修复

## 许可证

提交代码即表示您同意您的贡献将按照项目的 MIT 许可证进行许可。

## 联系方式

如有问题，请通过以下方式联系：

- GitHub Issues: https://github.com/luc-yang/DocxLib/issues
- Email: your.email@example.com

---

再次感谢您的贡献！
