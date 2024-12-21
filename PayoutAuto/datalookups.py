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