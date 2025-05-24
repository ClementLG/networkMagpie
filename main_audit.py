# main_audit.py
# !/usr/bin/env python3

# NetworkMagpie
# Copyright (C) 2025 CLEMENT LE GRUIEC
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import os
import datetime

# Importer les fonctions principales des scripts d'audit
# et les fonctions de chargement génériques (si elles ne sont pas déjà dans les scripts d'audit)
from cisco_audit import perform_cisco_audit, load_inventory as load_inventory_cisco, \
    load_passwords as load_passwords_cisco, generate_excel_report as generate_excel_cisco
from aruba_audit import perform_aruba_audit, load_inventory as load_inventory_aruba, \
    load_passwords as load_passwords_aruba, generate_excel_report as generate_excel_aruba


# Note: Les fonctions load_inventory et load_passwords sont identiques. On peut en choisir une version.
# Pour la simplicité, et si elles sont bien dans chaque script, on peut les appeler ainsi.
# Ou mieux : créer un fichier utils.py avec ces fonctions partagées.

def main():
    inventory_file = "inventory.csv"
    password_file = "passwords.csv"
    base_output_directory = "audit_reports"  # Répertoire de base pour tous les audits

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    # Créer un sous-répertoire horodaté pour cette session d'audit
    # session_output_directory = os.path.join(base_output_directory, f"audit_session_{timestamp}")
    # Ou garder un seul répertoire pour les rapports, avec des noms de fichiers horodatés (plus simple)
    session_output_directory = base_output_directory

    if not os.path.exists(session_output_directory):
        try:
            os.makedirs(session_output_directory)
        except OSError as e:
            print(f"Erreur: Impossible de créer le répertoire de sortie '{session_output_directory}': {e}")
            return

    # Utiliser une des fonctions load_inventory (elles devraient être identiques et lire device_type)
    # Assurez-vous que la fonction load_inventory dans cisco_audit.py et aruba_audit.py
    # lit bien la 3ème colonne 'device_type'.
    full_inventory = load_inventory_cisco(inventory_file)  # ou load_inventory_aruba
    passwords_map = load_passwords_cisco(password_file)  # ou load_passwords_aruba

    if full_inventory is None or passwords_map is None:
        print("Erreur critique lors du chargement des fichiers d'inventaire ou de mots de passe. Arrêt.")
        return

    if not full_inventory:
        print("L'inventaire est vide. Rien à faire.")
        return

    cisco_devices_to_audit = []
    aruba_devices_to_audit = []

    for device in full_inventory:
        dev_type = device.get("device_type", "").lower()  # Normaliser en minuscules
        if dev_type == "cisco_ios" or dev_type == "cisco_iosxe":  # Accepter les deux pour Cisco
            cisco_devices_to_audit.append(device)
        elif dev_type == "aruba_os-cx":
            aruba_devices_to_audit.append(device)
        else:
            print(
                f"Avertissement: Type d'équipement inconnu ou manquant '{device.get('device_type')}' pour {device.get('host')}. Ignoré.")

    if cisco_devices_to_audit:
        print(f"\n--- Lancement de l'audit pour {len(cisco_devices_to_audit)} équipement(s) Cisco ---")
        perform_cisco_audit(cisco_devices_to_audit, passwords_map, session_output_directory)
    else:
        print("\n--- Aucun équipement Cisco à auditer ---")

    if aruba_devices_to_audit:
        print(f"\n--- Lancement de l'audit pour {len(aruba_devices_to_audit)} équipement(s) Aruba ---")
        perform_aruba_audit(aruba_devices_to_audit, passwords_map, session_output_directory)
    else:
        print("\n--- Aucun équipement Aruba à auditer ---")

    print("\n\nTous les audits planifiés sont terminés.")


if __name__ == "__main__":
    main()
