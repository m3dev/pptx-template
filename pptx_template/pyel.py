# PyEl - simple expression language for Python
# coding=utf-8
#
# eval_el( 'foo.var' , { foo: { bar: 'Hello' } })
# ==> 'Hello'
#
# eval_el('bar.0', {bar: [1,2])
# ==> 1
#
# eval_el('0', [1,2])
# ==> 1
#
# build_el( { foo: { bar: 'Hello' }, list: [0, 1] })
# ==> [('foo.bar','Hello'), ('list.0', 0), ('list.1', 1)]
#

def eval_el(el, model):
    """
        el で示された値を model の中から取り出す。見つからない場合は ValueError 例外となる
    """
    context = model
    path = ''

    for part in el.split('.'):
        path = path + '.' + part

        if isinstance(context, list):
            if not part.isdigit():
                raise ValueError("%s: %s should be array but %s" % (el, part, context))
            index = int(part)
            if index < 0 or index >= len(context):
                raise ValueError("%s: index %d is out of range for model: %s" % (el, index, context))

            context = context[index]
            continue

        if isinstance(context, dict):
            if not part in context:
                raise ValueError("%s: %s key not found in model: %s" %(el, part, context))

            context = context[part]
            continue

        ValueError("model doesn't match %s" % path)
        break

    return context

def _flatten(list):
    return [item for sublist in list for item in sublist]

def _build_el_recursive(obj, path):
    delimiter = '' if path == '' else '.'
    result = []
    if isinstance(obj, list):
        return _flatten([ _build_el_recursive(obj[i], "%s%s%d" % (path, delimiter, i)) for i in range(len(obj)) ])
    elif isinstance(obj, dict):
        return _flatten([ _build_el_recursive(obj[k], "%s%s%s" % (path, delimiter, k)) for k in obj.keys() ])
    else:
        return [ (path, obj) ]

def build_el(obj):
    """
    obj の中に含まれるすべての値を (el, 値) の tuple の配列に展開する
    """
    return _build_el_recursive(obj, '')



def set_value(model, el, value):
    """
    model に対して、EL で指定された位置に value の値を追加する。
    EL で指定されたパスが存在しない場合は生成する。
    model と EL が不整合だった場合は ValueError
    """
    context = model
    path = ''

    def gen_part(parts):
        if len(parts) == 0:
            raise ValueError(u"EL is empty")
        for i in range(0, len(parts)):
            if not parts[i]:
                raise ValueError(u"EL %s should not include empty part" % el)
            yield (parts[i], parts[i+1] if (i+1) < len(parts) else None)

    for part, child in gen_part(el.split('.')):
        path = "%s.%s" % (path, part)

        index = int(part) if part.isdigit() else part

        if child == None:
            context[index] = value
            break

        if child.isdigit():
            if index not in context:
                context[index] = [None for i in range(0, int(child) + 1)]
            elif not isinstance(context[index], list):
                raise ValueError("context not match: %s" % path)
            elif int(child) >= len(context[index]):
                context[index].extend([None for i in range(len(context[index]), int(child) + 1)])
        else:
            if index not in context:
                context[index] = {}
            elif not isinstance(context[index], dict):
                raise ValueError("context not match: %s" % path)

        context = context[index]

    return model
