# 安全说明

WordReplace 是一个自托管工具，默认面向个人或内部使用。它**没有内置账号密码体系**，因此部署方式直接决定了安全性。

> 返回 [README](README.md) · 部署见 [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) · 排障见 [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

---

## 1. 默认密钥警告（最重要）

代码中的默认导出密钥是：

```text
change-this-secret
```

使用默认值时，应用启动日志会出现警告：

```text
EXPORT_TOKEN_SECRET is using the default value.
Set a strong random secret before exposing this service publicly.
```

风险：

- 知道默认值的人可以伪造导出令牌。
- 如果服务暴露到公网，可能被未授权下载结果。

处理：

1. 生成强随机密钥：

```bash
openssl rand -hex 32
```

2. 写入 `docker-compose.yml`：

```yaml
    environment:
      EXPORT_TOKEN_SECRET: "把上一步生成的随机字符串粘贴到这里"
```

3. 重启：

```bash
docker compose up -d
```

4. 确认日志不再出现默认密钥警告：

```bash
docker compose logs --tail=100 wordreplace
```

> 修改密钥后，之前生成的导出链接会立即失效，这是预期行为。重要结果请提前下载。

---

## 2. ALLOWED_HOSTS

`ALLOWED_HOSTS` 控制允许访问应用的 Host 头。

- 默认值：`*`，表示不限制。
- 当值不是 `*` 时，应用会启用 `TrustedHostMiddleware`，拒绝不匹配的 Host。

公网部署建议：

```yaml
    environment:
      ALLOWED_HOSTS: "word.example.com"
```

多个域名用英文逗号分隔：

```yaml
      ALLOWED_HOSTS: "word.example.com,www.example.com"
```

注意：

- 填域名，不要带 `https://`，不要带路径。
- 如果通过反向代理访问，代理必须正确转发 `Host` 头。
- 本机调试可以保留 `*`，但不要长期用于公网。

---

## 3. HTTPS 与反向代理

强烈建议通过 HTTPS 访问，尤其是上传包含真实姓名、合同、证件等内容的文档时。

推荐做法：

- 使用 Caddy、Nginx、Traefik 或群晖反向代理终止 HTTPS。
- 只把容器端口绑定到本机，例如：

```yaml
    ports:
      - "127.0.0.1:12344:8000"
```

- 由反向代理对外提供 `443`，并把请求转发到 `127.0.0.1:12344`。
- 同时设置：

```yaml
    environment:
      ALLOWED_HOSTS: "word.example.com"
      CORS_ALLOW_ORIGINS: "https://word.example.com"
```

Caddy 与 Nginx 的完整示例见 [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)。

注意：

- 反向代理的请求体上限要大于等于 `MAX_UPLOAD_SIZE_MB`，否则大文件会被代理拦截。
- 不要把容器端口直接暴露到公网而不加任何保护。

---

## 4. 建议增加访问认证

本项目默认没有登录、用户、角色或权限控制。任何能访问页面的人都可以：

- 上传文件。
- 执行替换。
- 使用自己的 `run_id` 与 `token` 下载或删除本次结果。
- 查看、创建、修改规则模板。

如果必须在公网使用，建议在反向代理层增加认证，例如：

- Nginx Basic Auth。
- Caddy `basic_auth`。
- Authelia、Authentik、Keycloak 等统一认证。
- 仅允许 VPN / 内网 IP 访问。
- 使用 Cloudflare Access 等零信任网关。

在反向代理层增加认证的示意（Nginx）：

```nginx
    auth_basic "WordReplace";
    auth_basic_user_file /etc/nginx/.htpasswd;
```

生成密码文件：

```bash
sudo apt-get install apache2-utils
sudo htpasswd -c /etc/nginx/.htpasswd youruser
```

> 应用内没有权限隔离：通过认证的人可以看到同一实例内的模板并使用全部功能；具体结果仍受 `run_id` 与导出令牌保护。请按“单团队共用”来评估风险。

---

## 5. 上传限制

默认限制：

| 项目 | 默认值 | 环境变量 |
|---|---|---|
| 单文件大小 | 50MB | `MAX_UPLOAD_SIZE_MB` |
| Excel 行数 | 20000 | `MAX_EXCEL_ROWS` |
| Excel 单元格总数 | 200000 | `MAX_EXCEL_CELLS` |
| 单次规则数量 | 500 | `MAX_REPLACE_RULES` |
| Word 预览字符数 | 20000 | `MAX_PREVIEW_CHARS` |

应用会：

- 只接受 `.docx`、`.xlsx`、`.xls`。
- 检查文件头，拒绝改后缀的假文件。
- 在读取过程中实时累计大小，超过限制立即拒绝。
- 拒绝空文件。

建议：

- 不要在公网实例上盲目调大限制。
- 调大 `MAX_UPLOAD_SIZE_MB` 时，同步调整反向代理的 `client_max_body_size`。
- 如果只处理固定格式的小文档，保持默认或更低限制。

---

## 6. 结果缓存与过期

- 替换结果保存在**进程内存**中，不写入磁盘。
- 默认 30 分钟后过期（`RUN_CACHE_TTL_SECONDS`）。
- 缓存还有记录数上限与总字节上限。
- 容器重启会清空所有结果。
- 导出令牌带 HMAC 签名与过期时间，过期后无法下载。

建议：

- 在结果过期前下载并保存到安全位置。
- 不要把带 `token` 的下载链接发到公开群聊或工单系统。
- 不需要的结果可以点击删除，或等其自动过期。
- 公网使用时不要过度调大 TTL，减少结果驻留时间。

---

## 7. 不要在没有保护的情况下公开

以下做法**不安全**：

- 直接把 `12344` 端口映射到公网。
- 使用默认 `EXPORT_TOKEN_SECRET`。
- `ALLOWED_HOSTS` 设为 `*` 且没有反向代理认证。
- 用 HTTP 明文传输含个人信息的文档。
- 把 `.env`、compose 密钥、证书私钥提交到 Git。

最低安全配置建议：

```yaml
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:3.0.0
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "127.0.0.1:12344:8000"
    volumes:
      - ./data:/app/data
    environment:
      ALLOWED_HOSTS: "word.example.com"
      CORS_ALLOW_ORIGINS: "https://word.example.com"
      EXPORT_TOKEN_SECRET: "至少 32 位随机字符串"
      MAX_UPLOAD_SIZE_MB: "50"
      RUN_CACHE_TTL_SECONDS: "1800"
```

再加一层反向代理认证和 HTTPS。

---

## 8. 其他安全措施

- 容器以 UID `10001` 的非 root 用户运行。
- 应用返回基础安全响应头：
  - `X-Content-Type-Options: nosniff`
  - `X-Frame-Options: DENY`
  - `Referrer-Policy: no-referrer`
  - `Permissions-Policy`
  - 有前端资源时附加 CSP。
- CORS 只允许 `CORS_ALLOW_ORIGINS` 中配置的来源。
- 生成文件名会过滤路径分隔符、Windows 非法字符和保留名。
- 仓库提供 `.githooks/pre-commit`，可在提交前扫描常见密钥格式。

启用 Git 钩子：

```bash
git config core.hooksPath .githooks
```

密钥应通过环境变量或密钥管理工具注入，不要写进代码或提交到仓库。

---

## 9. 漏洞反馈

如果你发现安全问题，请不要公开提交包含利用细节的 Issue。

建议通过 GitHub 仓库的 **Security → Report a vulnerability** 私密报告：

```text
https://github.com/MaroD1M/WordReplace
```

请在报告中说明：

- 受影响版本。
- 复现步骤。
- 影响范围。
- 建议的修复方式（如有）。

---

## 10. 安全部署检查清单

- [ ] 已修改 `EXPORT_TOKEN_SECRET`，不再是 `change-this-secret`。
- [ ] `ALLOWED_HOSTS` 已设置为真实域名，不是 `*`。
- [ ] 已通过 HTTPS 访问。
- [ ] 容器端口只绑定到 `127.0.0.1` 或内网。
- [ ] 反向代理已增加认证，或服务仅限内网 / VPN。
- [ ] `MAX_UPLOAD_SIZE_MB` 与代理请求体上限匹配。
- [ ] 已确认 `./data` 权限正确，敏感数据目录仅授权人员可访问。
- [ ] 已确认日志中没有默认密钥警告。
- [ ] 已启用 Git 钩子并避免提交密钥。
- [ ] 已定期备份数据库，并妥善保管备份文件。
