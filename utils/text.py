def normalize(text):
    cleaned = ''.join(c if c.isprintable() else ' ' for c in str(text))
    return ' '.join(cleaned.split())
