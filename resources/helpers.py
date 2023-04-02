def get_node_value(item, tag_name):
    if all([item.getElementsByTagName(tag_name),
            item.getElementsByTagName(tag_name)[0],
            item.getElementsByTagName(tag_name)[0].firstChild,]):
        return item.getElementsByTagName(tag_name)[0].firstChild.nodeValue if item.getElementsByTagName(tag_name)[0].firstChild else None
    return
