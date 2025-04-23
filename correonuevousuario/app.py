import os
import pythoncom
import win32com.client as win32
from flask_cors import CORS
from flask import Flask, request, render_template

app = Flask(__name__)
CORS(app)

@app.route('/')
def formulario():
    return render_template('formulario.html')

@app.route('/enviar', methods=['POST'])
def enviar_correo():
    print("📨 Formulario recibido")
    print(request.form)

    # Inicializar COM
    pythoncom.CoInitialize()

    # Datos del formulario
    nombre = request.form['nombre_completo']
    titulo = request.form['titulo']
    sexo = request.form['sexo']
    correo = request.form['correo']
    usuario = request.form['usuario']
    contrasena = request.form['contrasena']

    titulo_saludo = f"{titulo} {nombre}" if titulo else nombre
    saludo = f"¡Bienvenida {titulo_saludo}!" if sexo == "mujer" else f"¡Bienvenido {titulo_saludo}!"
    asunto = f"¡Bienvenida al equipo, {titulo_saludo}!" if sexo == "mujer" else f"¡Bienvenido al equipo, {titulo_saludo}!"

    cuerpo = f"""
    <p>{saludo}</p>
    <p>Estamos muy emocionados de que te unas a nuestro equipo.</p>
    <p>Para facilitarte el inicio de tus actividades, te compartimos tus credenciales de acceso a Microsoft 365:</p>
    <ul>
        <li><strong>Nombre para mostrar:</strong> {nombre}</li>
        <li><strong>Nombre de usuario:</strong> {usuario}</li>
        <li><strong>Contraseña:</strong> {contrasena}</li>
    </ul>
    <p>Por favor, accede a <a href="https://www.office.com">www.office.com</a> y actualiza tu contraseña en tu primer inicio de sesión.</p>
    """

    cuerpo2 = f"""
    <p>Estimado {nombre},</p>
    <p>Espero que te encuentres bien. A continuación, te proporciono información importante:</p>
    <ul>
        <li><strong>Acceso a Teams:</strong> Usa tu correo empresarial para ingresar.</li>
        <li><strong>Acceso a la aplicación Delihealths:</strong> Descárgala y usa tu correo registrado <strong>{usuario}</strong>.</li>
        <li><strong>Capacitación:</strong> Adjunto encontrarás tu firma de correo.</li>
    </ul>
    <p>Saludos.</p>
    """

    # firma_path = "firma_delihealths.docx"
    firma_path = r"C:\Users\oigre\Desktop\altausuario\correonuevousuario\firma_delihealths.docx"    
    # firma_absoluta = os.path.abspath(firma_path)

    # Intentar enviar los correos
    outlook = None
    try:
        print("🛠️ Inicializando Outlook...")
        outlook = win32.Dispatch("Outlook.Application")

        print("📧 Enviando primer correo...")
        mail1 = outlook.CreateItem(0)
        mail1.To = correo
        mail1.Subject = asunto
        mail1.HTMLBody = cuerpo
        mail1.Send()
        print(f"✅ Correo de bienvenida enviado a {correo}")

        # Cerrar Outlook después del primer correo
        outlook.Quit()
        print("🔒 Outlook cerrado después del primer correo.")

    except Exception as e:
        print(f"❌ Error al enviar el primer correo: {e}")

    try:
        print("📎 Preparando segundo correo con firma...")
        if outlook:
            outlook = win32.Dispatch("Outlook.Application")  # Reiniciar Outlook para el segundo correo
            mail2 = outlook.CreateItem(0)
            mail2.To = usuario
            mail2.Subject = "Accesos y capacitación - Grupo Delihealths"
            mail2.HTMLBody = cuerpo2
            mail2.Attachments.Add(firma_path)
            # mail2.Attachments.Add(firma_absoluta)  # Si necesitas la ruta absoluta, descomentar esta línea    
            mail2.Send()
            print(f"✅ Correo con firma enviado a {usuario}")

            # Cerrar Outlook después del segundo correo
            outlook.Quit()
            print("🔒 Outlook cerrado después del segundo correo.")
        else:
            print("⚠️ Outlook no disponible. No se puede enviar el segundo correo.")
    except Exception as e:
        print(f"❌ Error al enviar el segundo correo: {e}")

    return render_template("resultado.html", nombre=nombre)

if __name__ == '__main__':
    app.run(debug=True)
