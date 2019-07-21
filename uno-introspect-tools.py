import uno
from inspect import *

# this is a modified version of inspect.getmembers(). It was modified to not fail on
# uno exceptions that sometimes happen during access attempts
# And while on it: ignore uno.ByteSequence by default.
# returns: [(String, a)]
def getmembers_uno(object, predicate=lambda obj: not isinstance(obj, uno.ByteSequence)):
    """Return all members of an object as (name, value) pairs sorted by name.
    Optionally, only return members that satisfy a given predicate."""
    if isclass(object):
        mro = (object,) + getmro(object)
    else:
        mro = ()
    results = []
    processed = set()
    names = dir(object)
    # :dd any DynamicClassAttributes to the list of names if object is a class;
    # this may result in duplicate entries if, for example, a virtual
    # attribute with the same name as a DynamicClassAttribute exists
    try:
        for base in object.__bases__:
            for k, v in base.__dict__.items():
                if isinstance(v, types.DynamicClassAttribute):
                    names.append(k)
    except AttributeError:
        pass
    for key in names:
        # First try to get the value via getattr.  Some descriptors don't
        # like calling their __get__ (see bug #1785), so fall back to
        # looking in the __dict__.
        try:
            value = getattr(object, key)
            # handle the duplicate key
            if key in processed:
                raise AttributeError
        except AttributeError:
            for base in mro:
                if key in base.__dict__:
                    value = base.__dict__[key]
                    break
            else:
                # could be a (currently) missing slot member, or a buggy
                # __dir__; discard and move on
                continue
        except uno.getClass("com.sun.star.uno.RuntimeException"):
            continue # ignore: inspect.RuntimeException: Getting from this property is not supported
        except Exception:
            continue # ignore: everything, we don't care
        if not predicate or predicate(value):
            results.append((key, value))
        processed.add(key)
    results.sort(key=lambda pair: pair[0])
    return results

def isiter(maybe_iterable):
    try:
        iter(maybe_iterable)
        return True
    except TypeError:
        return False

def getValSafe(obj, prop_s):
    try:
        return getattr(obj, prop_s)
    except Exception:
        return None

def getmembers_uno2(object,
                     predicate=lambda obj: not isinstance(obj, uno.ByteSequence)):
    ret = []
    for prop_s in dir(object):
        val = getValSafe(object, prop_s)
        if predicate(val):
            ret.append((prop_s, val))
    return ret

# nLevels: number of levels it allowed to descend
# predicate: types to use, except iterables
# pIgnore: String -> Bool, accepts property name, may be used to drive search.
def searchLimited(unoObject, valToSearch, nLevels,
                  predicate, pIgnore, path = '<TOPLEVEL>'):
    def try_property(property_name, val):
        if property_name.startswith('__') or pIgnore(property_name):
            return None # ignore private stuff
        print('TRACE: ' + path + '.' + property_name)
        if val == valToSearch:
            return path + '.' + property_name
        if nLevels - 1 != 0 and isiter(val):
            index = 0
            for item in val:
                ret = searchLimited(item, valToSearch, nLevels - 1, predicate, pIgnore,
                                    path + '.' + property_name + '.<iter ' + str(index) + '>')
                if (ret != None):
                    return ret
                index += 1
    for (property_name, val) in getmembers_uno(unoObject, lambda p: isiter(p) or predicate(p)):
        ret = try_property(property_name, val)
        if (ret != None):
            return ret
    if nLevels - 1 != 0 and isiter(unoObject):
        index = 0
        for item in unoObject:
            ret = searchLimited(item, valToSearch, nLevels - 1, predicate, pIgnore,
                                 path + '.<iter ' + str(index) + '>')
            if (ret != None):
                return ret
            index += 1
    return None
