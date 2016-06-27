#!/usr/bin/perl -w
binmode(STDOUT, ":utf8");
binmode(STDERR, ":utf8");
use utf8;
use strict;
use encoding 'utf8';

main:
{
	while(<>) {
		s/፡፡/።/g;
		s/ኀ\+/ኄ/g;
		s/ኄ\+/ኄ/g;
		s/ÿኄ/ኌ/g;
		s/È/»/g;
		# s/È/"/g;
		print;
		# ÿኀ+    ኍ
		# ÿካ+    ከ=  
		# ÿቄ    ከ=
	}

}
