from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
import time
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
import os
from PIL import Image
import io
import zipfile

@task
def order_robots_from_RobotSpareBin():
 browser.configure(slowmo=100)
 ouvrir_site_robot()
 fermer_modal_enervant()
 time.sleep(10)
 telecharger_fichier_csv()
 csv()
 zipper_les_pdfs()

def ouvrir_site_robot():
 browser.goto("https://robotsparebinindustries.com/#/robot-order")

def fermer_modal_enervant():
    """Ferme la fenêtre modale ennuyeuse si elle apparaît"""
    try:
        page = browser.page()
        page.click("text=OK")
    except Exception as e:
        print(f"Aucune modal trouvée ou erreur survenue : {e}")

def telecharger_fichier_csv():
    """Télécharge le fichier CSV depuis l'URL donnée"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def csv():
    library = Tables()
    orders = library.read_table_from_csv("orders.csv")
    for row in orders:
        order_number = row['Order number']
        head = row['Head']
        body = row['Body']
        legs = row['Legs']
        address = row['Address']

        fermer_modal_enervant
        remplir_et_envoyer_formulaire_vente(body, head, legs, address)
        
        path_to_pdf = store_receipt_as_pdf(order_number,address)
        
        print(f"PDF enregistré à : {path_to_pdf}")

def remplir_et_envoyer_formulaire_vente(body, head, legs, address):
    """Remplit les données de vente et clique sur le bouton 'Commander'"""
    page = browser.page()

    # Attendre que l'élément #head soit présent 
    try:
        page.wait_for_selector("#head", timeout=4000) 
        page.select_option("#head", head)
        page.click(f"id=id-body-{body}")
        page.fill("//input[starts-with(@id, '172')]", legs)
        page.fill("#address", address)
        page.click("id=order")
        #time.sleep(20) 
        page.click("text=preview")
        browser.wait_for_selector("#robot-preview-image")
        page.click('#order')

        while page.query_selector('div.alert.alert-danger'):
            print("Erreur détectée : 'Server In Flames Error'. Tentative de clic pour résoudre le problème.")
            page.click("id=order")  
            time.sleep(2)

        

        browser.wait_for_selector("#receipt")
        print("Formulaire soumis avec succès.")
    except Exception as e:
        print(f"Erreur lors de la sélection de l'option de la tête : {e}")

    while page.query_selector('#order') :
        page.click("id=order") 
        time.sleep(2)
    return




def store_receipt_as_pdf(order_number,address):
    """Enregistre un PDF avec un nom de fichier unique basé sur le numéro de commande et inclut une capture d'écran"""
    pdf = PDF()
    fs = FileSystem()
    output_dir = os.path.join(os.getcwd(), "output", "receipts")
    os.makedirs(output_dir, exist_ok=True)  # Crée le répertoire s'il n'existe pas
    file_name = f"order_{order_number}.pdf"
    file_path = os.path.join(output_dir, file_name)
    
    # Capture d'écran de l'élément spécifique
    screenshot_path = os.path.join(output_dir, f"order_{order_number}_screenshot.png")
    page = browser.page()
    
    try:
        page.wait_for_selector('#receipt')
        page.wait_for_selector('#robot-preview-image')
        element1 = page.query_selector('#receipt h3')
        element3 = page.query_selector('#receipt div')
        element4 = page.query_selector('#receipt p.badge.badge-success')
        element_p = page.query_selector('#receipt p')
        element_parts = page.query_selector('#parts')
        element2 = page.query_selector('#robot-preview-image')

        if element1 and element2:
            # Extraire le texte de l'élément1
            texte_element1 = element1.text_content()
            texte_element3 = element3.text_content()
            texte_element4 = element4.text_content()
            texte_element_p = element_p.text_content()
            texte_element_parts = element_parts.text_content()

            
            # Capturer la capture d'écran de l'élément2
            screenshot_element2 = element2.screenshot()
            image2 = Image.open(io.BytesIO(screenshot_element2))
            image2.save(screenshot_path)

            # Créer le contenu du PDF
            pdf_content = f"""
            <html>
            <body>
            <h1>{texte_element1}</h1>
            <p>{texte_element3}</p>
            <p>{texte_element4}</p>
            <p>{address} {texte_element_parts}</p>
            <p>Thank you for your order! We will ship your robot to you as soon as our warehouse robots gather the parts you ordered! You will receive your robot in no time!</p>
            <p> </p>
            <p> </p>
            <div style="display: flex; justify-content: center; align-items: center; height: calc(100vh - 200px);">
            <img src="{screenshot_path}" alt="Capture d'écran de la commande" style="max-width: 90%; max-height: 90%; height: auto;">
            </div>
            </body>
            </html>
            """
            pdf.html_to_pdf(pdf_content, file_path)
        else:
            raise Exception("Les éléments 'receipt' ou 'robot-preview-image' n'ont pas été trouvés.")
    except Exception as e:
        print(f"Erreur lors de la capture des éléments : {e}")
    

    page.click("id=order-another")
    page = browser.page()
    page.click("text=OK")

    return file_path
    

def zipper_les_pdfs():
    """Compresse tous les PDF dans un fichier ZIP"""
    output_dir = os.path.join(os.getcwd(), "output", "receipts")
    zip_file_path = os.path.join(os.getcwd(), "output", "receipts.zip")
    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                if file.endswith('.pdf'):
                    zipf.write(os.path.join(root, file), arcname=file)
                    print(f"Tous les PDF ont été compressés dans {zip_file_path}")