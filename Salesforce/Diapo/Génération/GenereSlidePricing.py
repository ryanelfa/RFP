import subprocess

def main():
    presentation_id_source = '1FgHNfpr_I8Lf0wEQODghZGzsr6nN_LKiIt1kxG0F5hM'
    presentation_id_destination = '1FgHNfpr_I8Lf0wEQODghZGzsr6nN_LKiIt1kxG0F5hM'
    slide_id_source = 'g2ec28240e47_3_2582'

    result1 = subprocess.run(['python3', 'GenereTemplate.py', presentation_id_source, presentation_id_destination, slide_id_source], capture_output=True, text=True)
    print("Résultat de GenereTemplate.py :")
    idDiapoDestination = result1.stdout.strip()
    print(result1.stdout)
    print(result1.stderr)

    result2 = subprocess.run(['python3', 'CreeZoneTxtPricing.py', presentation_id_destination, idDiapoDestination], capture_output=True, text=True)
    print("Résultat de CreeZoneTxtPricing.py :")
    print(result2.stdout)
    print(result2.stderr)
if __name__ == "__main__":
    main()
