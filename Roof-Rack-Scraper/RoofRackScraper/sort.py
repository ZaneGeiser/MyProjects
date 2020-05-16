def shell_sort(list):
    """Sorts the given list in place"""
    n = int(len(list)/2)
    while n > 0:
        for i in range(len(list)-n):
            if list[i] > list[i+n]:
                temp = list[i+n]
                list[i+n] = list[i]
                list[i] = temp
        n -= 1
    pass

def find_part_num(string):
    """Sorts the given string and returns the longest number or alphanumeric word in the string.
    Functionaly used to find the part number from a list representation of a given string."""
    if string is None or string == "":
        return ""

    list = string.split()
    #create a new list of words that contain integers.
    new_list = []
    for word in list:
        if 'mm' in word or 'cm' in word: #exceptions to the longest word with integer rule as lengths will not be the part number.
            continue
        for char in word:
            try:
                int(char)
                new_list += [word]
            except:
                continue
    #if there are no integers, assume that there is not a part number.
    if len(new_list) == 0:
        return ""

    #remove the special characters in characters_to_remove.
    characters_to_remove = "()"
    for word in new_list:
        for character in characters_to_remove:
            word = word.replace(character, "")

    #return the longest word in the new list as this is almost always the part number.
    return max(new_list, key=len)
