import json
import argparse


def findBilligsteLevCombo(package):
    billigste_lev_combo = []
    billigste_lev_combo_pris = 0

    lev_obj_list = package['lev_obj_list']

    length_of_lev_obj_list = len(lev_obj_list)
    length_of_pricelist = len(lev_obj_list[0]['pris_liste'])

    lev_obj_list_idx_counter = 0
    for i in range(0, length_of_pricelist):
        billigste_lev = "none"
        current_cheapest_price = 99999999
        for j in range(0, length_of_lev_obj_list):
            if lev_obj_list[j]['pris_liste'][i] != 0 and lev_obj_list[j]['pris_liste'][i] < current_cheapest_price:
                current_cheapest_price = lev_obj_list[j]['pris_liste'][i]
                billigste_lev = lev_obj_list[j]['lev_navn']
        billigste_lev_combo_pris += current_cheapest_price
        billigste_lev_combo.append(billigste_lev)

    return [billigste_lev_combo, billigste_lev_combo_pris]

def conflictFinder(package):    
    current_lev = ''
    for idx, lev in enumerate(package['billigste_lev_combo']):
        if idx == 0:
            current_lev = lev
        elif current_lev != lev:
            return [True, package]
    return [False, {}]


#Command Line Argument Parser
parser = argparse.ArgumentParser(description = "Specify input-/output- file/s")
parser.add_argument(  '-i', default='pakker_json.json')
args = parser.parse_args()

f = open('pakker_incorrect.json','rb')
root = json.load(f)

pack_list = root['pakker']

root2 = {"conflicting_packages":[]}
conflicting_packages = root2['conflicting_packages']

for package in pack_list:
    theCombo = findBilligsteLevCombo(package)
    package['billigste_lev_combo'] = theCombo[0]
    package['billigste_lev_combo_pris'] = theCombo[1]
    
    isConflict = conflictFinder(package)
    if isConflict[0] == True:
        conflicting_packages.append(isConflict[1])

f.close()


f = open('pakker_incorrect.json','wb')
f.write(json.dumps(root, indent=4))
f.close()

f = open('pakker_conflict.json','wb')
f.write(json.dumps(root2, indent=4))
f.close()