RAKU     := raku
# note LIBPATH uses normal PERL6LIB Perl 6 separators (',')
LIBPATH   := lib

# set below to 1 for no effect, 1 for debugging messages
DEBUG := MyMODULE_DEBUG=0

# set below to 0 for no effect, 1 to die on first failure
EARLYFAIL := RAKU_TEST_DIE_ON_FAIL=0


.PHONY: test bad good

default: test

TESTS     := t/*.t
BADTESTS  := bad-tests/*.t
GOODTESTS := good-tests/*.t

# the original test suite (i.e., 'make test')
test:
	for f in $(TESTS) ; do \
	    $(DEBUG) $(EARLYFAIL) PERL6LIB=$(LIBPATH) prove -v --exec=$(RAKU) $$f ; \
	done

bad:
	for f in $(BADTESTS) ; do \
	    $(DEBUG) $(EARLYFAIL) PERL6LIB=$(LIBPATH) prove -v --exec=$(RAKU) $$f ; \
	done

good:
	for f in $(GOODTESTS) ; do \
	    $(DEBUG) $(TA) $(EARLYFAIL) PERL6LIB=$(LIBPATH) prove -v --exec=$(RAKU) $$f ; \
	done
