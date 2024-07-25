import subprocess

def main():
    presentation_id_destination = '1FgHNfpr_I8Lf0wEQODghZGzsr6nN_LKiIt1kxG0F5hM'
    idDiapoDestination = 'g2eb0071ea72_0_3189'
    result2 = subprocess.run(['python3', 'CreeZoneTxtDeclarationConditionnelles.py', presentation_id_destination, idDiapoDestination], capture_output=True, text=True)
    print("Résultat de CreeZoneTxtDeclarationConditionnelles.py :")
    print(result2.stdout)
    print(result2.stderr)
if __name__ == "__main__":
    main()
