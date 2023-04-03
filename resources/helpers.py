from datetime import datetime


def get_node_value(item, tag_name):
    if all([item.getElementsByTagName(tag_name),
            item.getElementsByTagName(tag_name)[0],
            item.getElementsByTagName(tag_name)[0].firstChild,]):
        return item.getElementsByTagName(tag_name)[0].firstChild.nodeValue if item.getElementsByTagName(tag_name)[0].firstChild else None
    return


def to_chunks(xs, n):
    """ Split a list into list of chunks of length of n """
    n = max(1, n)
    return (xs[i:i+n] for i in range(0, len(xs), n))


def create_dt(datetime_string, datetime_format):
    dt = None 

    if datetime_string:
        try:
            dt = datetime.strptime(datetime_string, datetime_format)
        except Exception as e:
            print("ERROR WITH DATETIME STRING", datetime_string, "===>", str(e))

    return dt