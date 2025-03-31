def artist_lookup(artist_name, artist_info):
    '''
    This lookup does a more in-depth search for an artist if they were
    not found immediately.

    @param artist_name: The name of the artist we are looking for
    @param artist_info: The json containing all artist information

    @return The key of the artist if found, None otherwise
    '''

    for key, value in artist_info.items():

        if artist_name in value["aliases"] or artist_name == value["name"]:
            return key
        
    return None


def identify_vendors(product_vendor, second_product_vendor, artist_info):
    '''
    This lookup identifies the role of the product vendor and the second_product_vendor for an item.
    This way we know which is the artist and which is the collaborator. This is important as we need this information
    to determine splits
    '''
    if (product_vendor == "Saa Halust" and second_product_vendor == "Rozen") or (product_vendor == "Rozen" and second_product_vendor == "Saa Halust"):
        print("TEST")

    #vendor indentification
    product_vendor_role = artist_info[product_vendor]["client_type"]
    second_product_vendor_role = artist_info[second_product_vendor]["client_type"]

    #check validity of vendor identification - sceneario where both have only a singular role (most common)
    if ("Artist" in product_vendor_role and len(product_vendor_role) == 1) and ("Collab" in second_product_vendor_role and len(product_vendor_role) == 1):
        return {"Artist": product_vendor, "Collab": second_product_vendor}
    
    elif ("Artist" in second_product_vendor_role and len(second_product_vendor_role) == 1) and ("Collab" in product_vendor_role and len(product_vendor_role) == 1):
        return {"Artist": second_product_vendor, "Collab": product_vendor}
    
    #If one has multiple roles, we should check and see if the other has the collab role to differentiate
    elif "Artist" in product_vendor_role and "Artist" in second_product_vendor_role:

        if "Collab" in product_vendor_role and "Collab" not in second_product_vendor_role:
            return {"Artist": product_vendor, "Collab": second_product_vendor}
        
        if "Collab" in second_product_vendor_role and "Collab" not in product_vendor_role:
            return {"Artist": second_product_vendor, "Collab": product_vendor}
        
        else:
            print("ERROR:", product_vendor, "and", second_product_vendor, "both are of the artist type. Please check the artist information and correct the roles or correct the item manually.")
            
            #This is a breaking error and should exit the program
            exit()

    elif "Collab" in product_vendor_role and "Collab" in second_product_vendor_role:

        if "Artist" in product_vendor_role and "Artist" not in second_product_vendor_role:
            return {"Artist": product_vendor, "Collab": second_product_vendor}
        
        if "Artist" in second_product_vendor_role and "Artist" not in product_vendor_role:
            return {"Artist": second_product_vendor, "Collab": product_vendor}

        print("ERROR:", product_vendor, "and", second_product_vendor, "both are of the collab type. Please check the artist information and correct the roles or correct the item manually.")
            
        #This is a breaking error and should exit the program
        exit()

    return None


def isPremium(product_vendor, artist_info): 
    
    if artist_info[product_vendor]["profit_split_premium"]:
        return True
    
    return False