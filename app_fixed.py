def norm_q(v):
    s = clean(v).upper()
    compact = ''.join(ch for ch in s if ch.isalnum())
    if 'SCOPUSQ1' in compact:
        return 'Q1'
    if 'SCOPUSQ2' in compact:
        return 'Q2'
    return ''
