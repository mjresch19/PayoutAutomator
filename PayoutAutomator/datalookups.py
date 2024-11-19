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

def item_lookup(item_name, item_info):
    '''
    This function looks up an item that was labeled as YNM Vended. If it exists
    then we return the artists, if not then we return None.

    @param item_name: The name of the item we are looking for
    @param item_info: The json containing all item information

    @return The artist name associated with the item
    '''
    return item_info.get(item_name, None)


def name_extract(product_name, artist_info):
    '''
    This function will work to extract the collaborator's name from our artist info dictionary

    @param product_name: The name of the product we are looking at
    @param artist_info: The json containing all artist information

    @return The name of the collaborator, if we cannot find it, we consider this an imporperly named item and return None

    '''

    #Alias search the item to collab name
    for key, value in artist_info.items():

        for val in value["aliases"]:

            if val.lower() in product_name.lower():

                return key

    return None