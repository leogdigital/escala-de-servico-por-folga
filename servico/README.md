# Rodar como serviço no Linux

O app é um servidor Python. Para deixá-lo no ar depois do boot, use o systemd.

1. Copie a pasta do projeto para o servidor, por exemplo em `/caminho/do/app`.
2. Edite `escala.service.example` e troque:
   - `SEU_USUARIO` pelo usuário que vai rodar o app;
   - `/caminho/do/app` pelo caminho real;
   - `PERMITIR` pelos IPs que podem abrir a página.
3. Copie o arquivo para o systemd:

   ```bash
   sudo cp escala.service.example /etc/systemd/system/escala.service
   sudo systemctl daemon-reload
   sudo systemctl enable --now escala
   ```

4. Confira:

   ```bash
   systemctl status escala
   curl -s -o /dev/null -w "%{http_code}\n" http://127.0.0.1:8770/
   ```

Para acessar de outro computador da rede, troque no `ExecStart`:

```text
--host 127.0.0.1
```

por:

```text
--host 0.0.0.0
```

e coloque os IPs dos clientes em `PERMITIR`. Sem isso, o app só responde no
próprio servidor.

Para parar ou reiniciar:

```bash
sudo systemctl stop escala
sudo systemctl restart escala
```
