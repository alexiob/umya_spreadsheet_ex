{ lib, beamPackages, overrides ? (x: y: {}) }:

let
  buildRebar3 = lib.makeOverridable beamPackages.buildRebar3;
  buildMix = lib.makeOverridable beamPackages.buildMix;
  buildErlangMk = lib.makeOverridable beamPackages.buildErlangMk;

  self = packages // (overrides self packages);

  packages = with beamPackages; with self; {
    castore = buildMix rec {
      name = "castore";
      version = "1.0.14";

      src = fetchHex {
        pkg = "castore";
        version = "${version}";
        sha256 = "7bc1b65249d31701393edaaac18ec8398d8974d52c647b7904d01b964137b9f4";
      };

      beamDeps = [];
    };

    earmark_parser = buildMix rec {
      name = "earmark_parser";
      version = "1.4.44";

      src = fetchHex {
        pkg = "earmark_parser";
        version = "${version}";
        sha256 = "4778ac752b4701a5599215f7030989c989ffdc4f6df457c5f36938cc2d2a2750";
      };

      beamDeps = [];
    };

    ex_doc = buildMix rec {
      name = "ex_doc";
      version = "0.38.1";

      src = fetchHex {
        pkg = "ex_doc";
        version = "${version}";
        sha256 = "754636236d191b895e1e4de2ebb504c057fe1995fdfdd92e9d75c4b05633008b";
      };

      beamDeps = [ earmark_parser makeup_elixir makeup_erlang ];
    };

    finch = buildMix rec {
      name = "finch";
      version = "0.19.0";

      src = fetchHex {
        pkg = "finch";
        version = "${version}";
        sha256 = "fc5324ce209125d1e2fa0fcd2634601c52a787aff1cd33ee833664a5af4ea2b6";
      };

      beamDeps = [ mime mint nimble_options nimble_pool telemetry ];
    };

    glob_ex = buildMix rec {
      name = "glob_ex";
      version = "0.1.11";

      src = fetchHex {
        pkg = "glob_ex";
        version = "${version}";
        sha256 = "342729363056e3145e61766b416769984c329e4378f1d558b63e341020525de4";
      };

      beamDeps = [];
    };

    hpax = buildMix rec {
      name = "hpax";
      version = "1.0.3";

      src = fetchHex {
        pkg = "hpax";
        version = "${version}";
        sha256 = "8eab6e1cfa8d5918c2ce4ba43588e894af35dbd8e91e6e55c817bca5847df34a";
      };

      beamDeps = [];
    };

    igniter = buildMix rec {
      name = "igniter";
      version = "0.6.2";

      src = fetchHex {
        pkg = "igniter";
        version = "${version}";
        sha256 = "480c679066444459eecd371a41639cee365449e20cfb611457081394fd146b05";
      };

      beamDeps = [ glob_ex jason owl req rewrite sourceror spitfire ];
    };

    jason = buildMix rec {
      name = "jason";
      version = "1.4.4";

      src = fetchHex {
        pkg = "jason";
        version = "${version}";
        sha256 = "c5eb0cab91f094599f94d55bc63409236a8ec69a21a67814529e8d5f6cc90b3b";
      };

      beamDeps = [];
    };

    makeup = buildMix rec {
      name = "makeup";
      version = "1.2.1";

      src = fetchHex {
        pkg = "makeup";
        version = "${version}";
        sha256 = "d36484867b0bae0fea568d10131197a4c2e47056a6fbe84922bf6ba71c8d17ce";
      };

      beamDeps = [ nimble_parsec ];
    };

    makeup_elixir = buildMix rec {
      name = "makeup_elixir";
      version = "1.0.1";

      src = fetchHex {
        pkg = "makeup_elixir";
        version = "${version}";
        sha256 = "7284900d412a3e5cfd97fdaed4f5ed389b8f2b4cb49efc0eb3bd10e2febf9507";
      };

      beamDeps = [ makeup nimble_parsec ];
    };

    makeup_erlang = buildMix rec {
      name = "makeup_erlang";
      version = "1.0.2";

      src = fetchHex {
        pkg = "makeup_erlang";
        version = "${version}";
        sha256 = "af33ff7ef368d5893e4a267933e7744e46ce3cf1f61e2dccf53a111ed3aa3727";
      };

      beamDeps = [ makeup ];
    };

    meck = buildRebar3 rec {
      name = "meck";
      version = "0.9.2";

      src = fetchHex {
        pkg = "meck";
        version = "${version}";
        sha256 = "81344f561357dc40a8344afa53767c32669153355b626ea9fcbc8da6b3045826";
      };

      beamDeps = [];
    };

    mime = buildMix rec {
      name = "mime";
      version = "2.0.7";

      src = fetchHex {
        pkg = "mime";
        version = "${version}";
        sha256 = "6171188e399ee16023ffc5b76ce445eb6d9672e2e241d2df6050f3c771e80ccd";
      };

      beamDeps = [];
    };

    mint = buildMix rec {
      name = "mint";
      version = "1.7.1";

      src = fetchHex {
        pkg = "mint";
        version = "${version}";
        sha256 = "fceba0a4d0f24301ddee3024ae116df1c3f4bb7a563a731f45fdfeb9d39a231b";
      };

      beamDeps = [ castore hpax ];
    };

    nimble_options = buildMix rec {
      name = "nimble_options";
      version = "1.1.1";

      src = fetchHex {
        pkg = "nimble_options";
        version = "${version}";
        sha256 = "821b2470ca9442c4b6984882fe9bb0389371b8ddec4d45a9504f00a66f650b44";
      };

      beamDeps = [];
    };

    nimble_parsec = buildMix rec {
      name = "nimble_parsec";
      version = "1.4.2";

      src = fetchHex {
        pkg = "nimble_parsec";
        version = "${version}";
        sha256 = "4b21398942dda052b403bbe1da991ccd03a053668d147d53fb8c4e0efe09c973";
      };

      beamDeps = [];
    };

    nimble_pool = buildMix rec {
      name = "nimble_pool";
      version = "1.1.0";

      src = fetchHex {
        pkg = "nimble_pool";
        version = "${version}";
        sha256 = "af2e4e6b34197db81f7aad230c1118eac993acc0dae6bc83bac0126d4ae0813a";
      };

      beamDeps = [];
    };

    owl = buildMix rec {
      name = "owl";
      version = "0.12.2";

      src = fetchHex {
        pkg = "owl";
        version = "${version}";
        sha256 = "6398efa9e1fea70a04d24231e10dcd66c1ac1aa2da418d20ef5357ec61de2880";
      };

      beamDeps = [];
    };

    req = buildMix rec {
      name = "req";
      version = "0.5.10";

      src = fetchHex {
        pkg = "req";
        version = "${version}";
        sha256 = "8a604815743f8a2d3b5de0659fa3137fa4b1cffd636ecb69b30b2b9b2c2559be";
      };

      beamDeps = [ finch jason mime ];
    };

    rewrite = buildMix rec {
      name = "rewrite";
      version = "1.1.2";

      src = fetchHex {
        pkg = "rewrite";
        version = "${version}";
        sha256 = "7f8b94b1e3528d0a47b3e8b7bfeca559d2948a65fa7418a9ad7d7712703d39d4";
      };

      beamDeps = [ glob_ex sourceror text_diff ];
    };

    rustler = buildMix rec {
      name = "rustler";
      version = "0.36.1";

      src = fetchHex {
        pkg = "rustler";
        version = "${version}";
        sha256 = "f3fba4ad272970e0d1bc62972fc4a99809651e54a125c5242de9bad4574b2d02";
      };

      beamDeps = [ jason toml ];
    };

    rustler_precompiled = buildMix rec {
      name = "rustler_precompiled";
      version = "0.8.2";

      src = fetchHex {
        pkg = "rustler_precompiled";
        version = "${version}";
        sha256 = "63d1bd5f8e23096d1ff851839923162096364bac8656a4a3c00d1fff8e83ee0a";
      };

      beamDeps = [ castore rustler ];
    };

    sourceror = buildMix rec {
      name = "sourceror";
      version = "1.10.0";

      src = fetchHex {
        pkg = "sourceror";
        version = "${version}";
        sha256 = "29dbdfc92e04569c9d8e6efdc422fc1d815f4bd0055dc7c51b8800fb75c4b3f1";
      };

      beamDeps = [];
    };

    spitfire = buildMix rec {
      name = "spitfire";
      version = "0.2.0";

      src = fetchHex {
        pkg = "spitfire";
        version = "${version}";
        sha256 = "743daaee2d81a0d8095431729f478ce49b47ea8943c7d770de86704975cb7775";
      };

      beamDeps = [];
    };

    telemetry = buildRebar3 rec {
      name = "telemetry";
      version = "1.3.0";

      src = fetchHex {
        pkg = "telemetry";
        version = "${version}";
        sha256 = "7015fc8919dbe63764f4b4b87a95b7c0996bd539e0d499be6ec9d7f3875b79e6";
      };

      beamDeps = [];
    };

    text_diff = buildMix rec {
      name = "text_diff";
      version = "0.1.0";

      src = fetchHex {
        pkg = "text_diff";
        version = "${version}";
        sha256 = "d1ffaaecab338e49357b6daa82e435f877e0649041ace7755583a0ea3362dbd7";
      };

      beamDeps = [];
    };

    toml = buildMix rec {
      name = "toml";
      version = "0.7.0";

      src = fetchHex {
        pkg = "toml";
        version = "${version}";
        sha256 = "0690246a2478c1defd100b0c9b89b4ea280a22be9a7b313a8a058a2408a2fa70";
      };

      beamDeps = [];
    };
  };
in self

