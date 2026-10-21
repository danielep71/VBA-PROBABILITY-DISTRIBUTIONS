"""
Binary64 argument identity, shared by the grid writers and the
promotion/migration tools (#32). check_grid_coverage.py and audit_references.py
still keep their own copies.

The accuracy grids store each argument as a decimal INPUT TOKEN. The Excel
exporter parses it with Val(), so the value the library receives is the IEEE-754
binary64 Double that the token parses to - not the decimal written, and not the
design value the token was produced from. Three helpers express that once:

    double_key(text)        identity of a token: the 16-hex-digit bit pattern
                            of the Double it parses to ('' for a blank cell,
                            '?<text>' if it does not parse to a finite
                            Double). "0.85" and "0.84999999999999998" share
                            one key.
    as_reference_arg(x)     the exact Double an emitted argument denotes,
                            promoted back into mpmath, at which every
                            reference must be evaluated.
    format_double_arg(x)    the token emitted for an argument: 17 significant
                            digits of as_reference_arg(x), which parses back to
                            exactly that Double.

The canonical row key built on double_key lives in _grid_safety.canonical_key.
test_binary64.py proves the Python side of these rules over ordinary decimals,
17-digit round trips, the normal/subnormal boundary, the minimum subnormal,
one-ulp seams and the maximum finite Double. Parity of VBA Val() with Python's
float() on the same tokens can only be proven in Excel and remains open in #32.
"""
import math
import struct


def double_key(text):
    """IEEE-754 bit pattern of an argument cell; '' for blank, '?...' if unparsable.

    Python reads "1e999", "inf" and "nan" as non-finite floats; VBA Val() never
    produces those, so such a token is marked rather than given a key."""
    t = (text or "").strip()
    if t == "":
        return ""
    try:
        value = float(t)
    except ValueError:
        return "?" + t
    if not math.isfinite(value):
        return "?" + t
    return struct.pack(">d", value).hex()


def exact_binary64_mpf(x):
    """The exact binary64 value of x, promoted back into high precision.

    A Double is a dyadic rational with a 53-bit significand, so the quotient
    below is exact in mpmath's binary floating point at any working precision.
    """
    import mpmath as mp
    num, den = float(x).as_integer_ratio()
    return mp.mpf(num) / mp.mpf(den)


def as_reference_arg(x):
    """Canonicalise an emitted argument to the Double the exporter will receive.

    The committed argument field is an INPUT TOKEN, not an exact real number.
    The exporter does ParseDouble = Val(S), so the authoritative test input is
    the Double produced by Val() on the serialised token - not the Double
    nearest the pre-serialisation design value. Those two paths coincide almost
    everywhere, but they are a double rounding apart. Between 2^52 and 2^53 the
    Doubles are exactly the integers, so a 17-digit token ending in .5 is a tie
    that rounds half to even, while the unrounded design value it was printed
    from can round to the other neighbour:

        design 6235507341273914.5000001 -> nearest Double 6235507341273915
        token  "6235507341273914.5"     -> parse -> 6235507341273914 (even)

    reference_audit_post_canonicalisation.csv records the LogGamma argument
    "6235507341273914.5" (bits 4336272930122b3a) as such a token. Following the token reproduces the value the historical observation was
    measured at, which keeps the numeric key set invariant and needs no
    re-export. That Val() also rounds this tie to even is part of the Excel-side
    parser parity still open in #32.

    Python ints pass through: they are count inputs, exact in binary64, and
    downstream code treats them as ints. A string is already a token and parses
    straight to its Double. Lists and tuples are canonicalised element by
    element."""
    if isinstance(x, (list, tuple)):
        return type(x)(as_reference_arg(v) for v in x)
    if isinstance(x, int):
        return x
    if isinstance(x, str):
        return exact_binary64_mpf(x)
    import mpmath as mp
    return exact_binary64_mpf(mp.nstr(x, 17))


def format_double_arg(x):
    """The argument token emitted for x: 17 significant digits of its Double."""
    import mpmath as mp
    return mp.nstr(as_reference_arg(x), 17)
