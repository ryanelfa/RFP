import subprocess

def main():
    presentation_id_source = '1FgHNfpr_I8Lf0wEQODghZGzsr6nN_LKiIt1kxG0F5hM'
    presentation_id_destination = '1FgHNfpr_I8Lf0wEQODghZGzsr6nN_LKiIt1kxG0F5hM'
    slide_id_source = 'g2ee00bb1fa6_0_32'

    result1 = subprocess.run(['python3', 'GenereTemplate.py', presentation_id_source, presentation_id_destination, slide_id_source], capture_output=True, text=True)
    print("Résultat de GenereTemplate.py :")
    idDiapoDestination = result1.stdout.strip()
    print(result1.stdout)
    print(result1.stderr)

    # Vérifiez que idDiapoDestination n'est pas vide
    if not idDiapoDestination:
        print("Erreur : l'id de la diapositive de destination est vide.")
        return

    result2 = subprocess.run(['python3', './../Creation/CreeZoneTxtExecSummary.py', presentation_id_destination, idDiapoDestination], capture_output=True, text=True)
    print("Résultat de CreeZoneTxtExecSummary.py :")
    print(result2.stdout)
    print(result2.stderr)

if __name__ == "__main__":
    main()
