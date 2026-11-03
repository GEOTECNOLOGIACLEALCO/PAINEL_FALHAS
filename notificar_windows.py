import sys, subprocess

def notificar(titulo="Clealco Geotecnologia", mensagem="Base de Voos & Processamentos sincronizada com sucesso!"):
    ps_cmd = (
        "[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); "
        "$n = New-Object System.Windows.Forms.NotifyIcon; "
        "$n.Icon = [System.Drawing.SystemIcons]::Information; "
        f"$n.BalloonTipTitle = '{titulo}'; "
        f"$n.BalloonTipText = '{mensagem}'; "
        "$n.Visible = $true; "
        "$n.ShowBalloonTip(4000); "
        "Start-Sleep -Seconds 2; "
        "$n.Dispose()"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
            creationflags=0x08000000
        )
    except Exception as e:
        pass

if __name__ == "__main__":
    t = sys.argv[1] if len(sys.argv) > 1 else "Clealco Geotecnologia"
    m = sys.argv[2] if len(sys.argv) > 2 else "Base sincronizada com a nuvem com sucesso!"
    notificar(t, m)
