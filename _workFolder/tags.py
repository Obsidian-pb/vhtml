#%%
block_tags = [
    "a", "abbr", "b", "body", "button", "canvas", "dd", "div", "dl", "dt", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6",
    "head", "header", "i", "label", "li", "mark", "nav", "ol", "option", "p", "s", "script", "select", "small", "span", "strong",
    "style", "sub", "sup", "svg", "table", "tbody", "td", "textarea", "tfoot", "th", "thead", "time", "title", "tr", "ul", "u"
]

single_Tags = [
    "br", "img", "meta", "link", "hr", "img", "input"
]

print(";".join(sorted(block_tags)), end='\n\n')
print(";".join(sorted(single_Tags)))

# %%
