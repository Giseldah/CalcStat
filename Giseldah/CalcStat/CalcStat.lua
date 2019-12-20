_G.CalcStat = function(sname, slvl, sparam)

	local SN = string.upper(string.match(sname,"(%w+)"));

	local L = slvl;
	local N = 1;
	local C = "";

	if sparam ~= nil then
		if type(sparam) == "number" then
			N = sparam;
		else
			C = sparam;
		end
	end

	if SN < "PARTBLOCKPRATPC" then
		if SN < "EVADEPRATPCAPR" then
			if SN < "CRITDEFPRATPA" then
				if SN < "BPEPPRAT" then
					if SN < "BLOCKPBONUS" then
						if SN < "ARMOURLOWRAW" then
							if SN < "ADJTRAITPROGRATINGS" then
								if SN == "-VERSION" then
									return "1.3.0p";
								else
									return 0;
								end
							elseif SN > "ADJTRAITPROGRATINGS" then
								if SN == "ARMCATMP" then
									if L-DblCalcDev <= 0 then
										return 0;
									elseif L-DblCalcDev <= 1 then
										return 14877/12350;
									elseif L-DblCalcDev <= 2 then
										return 12415/12350;
									else
										return 9880/12350;
									end
								else
									return 0;
								end
							else
								if L-DblCalcDev <= 105 then
									return LinFmod(1,CalcStat("TraitProg",1,N),0.8*CalcStat("TraitProg",105,N),1,105,L)/CalcStat("TraitProg",L,N);
								elseif L-DblCalcDev <= 106 then
									return LinFmod(1,0.8*CalcStat("TraitProg",105,N),CalcStat("TraitProg",106,N),105,106,L)/CalcStat("TraitProg",L,N);
								elseif L-DblCalcDev <= 121 then
									return 1;
								elseif L-DblCalcDev <= 130 then
									return LinFmod(1,CalcStat("TraitProg",121,N),0.9*CalcStat("TraitProg",130,N),121,130,L)/CalcStat("TraitProg",L,N);
								else
									return 0.9;
								end
							end
						elseif SN > "ARMOURLOWRAW" then
							if SN < "ARMPROGMP" then
								if SN == "ARMPROG" then
									if L-DblCalcDev <= 400 then
										return CalcStat("ItemProg",L,N)/CalcStat("StdProg",75,N);
									elseif L-DblCalcDev <= 449 then
										return LinFmod(1,CalcStat("ItemProg",400,N)/CalcStat("StdProg",75,N),0.9*(CalcStat("ItemProg",449,N)/CalcStat("StdProg",75,N)),400,449,L);
									else
										return 0.9*(CalcStat("ItemProg",L,N)/CalcStat("StdProg",75,N));
									end
								else
									return 0;
								end
							elseif SN > "ARMPROGMP" then
								if SN > "ARMQTYLOWMP" then
									if SN == "ARMTYPEMP" then
										if L-DblCalcDev <= 0 then
											return 0;
										else
											return DataTableValue({9,9,18.75,30,15,25,12,37.5},L);
										end
									else
										return 0;
									end
								elseif SN == "ARMQTYLOWMP" then
									if L-DblCalcDev <= 0 then
										return 0;
									else
										return DataTableValue({50,50,50,50,50,49.8,50,50},L);
									end
								else
									return 0;
								end
							else
								if L-DblCalcDev <= 0 then
									return 0;
								elseif L-DblCalcDev <= 1 then
									return CalcStat("ArmProg",N,CalcStat("ProgBMitHeavy",L));
								elseif L-DblCalcDev <= 2 then
									return CalcStat("ArmProg",N,CalcStat("ProgBMitMedium",L));
								else
									return CalcStat("ArmProg",N,CalcStat("ProgBMitLight",L));
								end
							end
						else
							return CalcStat("ArmTypeMP",ArmCodeIndex(C,2))*CalcStat("ArmQtyLowMP",ArmCodeIndex(C,2))*CalcStat("ArmCatMP",ArmCodeIndex(C,1))*CalcStat("ArmProgMP",ArmCodeIndex(C,1),L);
						end
					elseif SN > "BLOCKPBONUS" then
						if SN < "BLOCKPRATPB" then
							if SN < "BLOCKPRATP" then
								if SN == "BLOCKPPRAT" then
									return CalcStat("BPEPPRat",L,N);
								else
									return 0;
								end
							elseif SN > "BLOCKPRATP" then
								if SN == "BLOCKPRATPA" then
									return CalcStat("BPEPRatPA",L);
								else
									return 0;
								end
							else
								return CalcStat("BPEPRatP",L,N);
							end
						elseif SN > "BLOCKPRATPB" then
							if SN < "BLOCKPRATPCAP" then
								if SN == "BLOCKPRATPC" then
									return CalcStat("BPEPRatPC",L);
								else
									return 0;
								end
							elseif SN > "BLOCKPRATPCAP" then
								if SN > "BLOCKPRATPCAPR" then
									if SN == "BPEPBONUS" then
										return 0;
									else
										return 0;
									end
								elseif SN == "BLOCKPRATPCAPR" then
									return CalcStat("BPEPRatPCapR",L);
								else
									return 0;
								end
							else
								return CalcStat("BPEPRatPCap",L);
							end
						else
							return CalcStat("BPEPRatPB",L);
						end
					else
						return CalcStat("BPEPBonus",L);
					end
				elseif SN > "BPEPPRAT" then
					if SN < "BRATMASTERY" then
						if SN < "BPEPRATPC" then
							if SN < "BPEPRATPA" then
								if SN == "BPEPRATP" then
									return CalcPercAB(CalcStat("BPEPRatPA",L),CalcStat("BPEPRatPB",L),CalcStat("BPEPRatPCap",L),N);
								else
									return 0;
								end
							elseif SN > "BPEPRATPA" then
								if SN == "BPEPRATPB" then
									return CalcStat("BratLow",L);
								else
									return 0;
								end
							else
								return 26;
							end
						elseif SN > "BPEPRATPC" then
							if SN < "BPEPRATPCAPR" then
								if SN == "BPEPRATPCAP" then
									return 13;
								else
									return 0;
								end
							elseif SN > "BPEPRATPCAPR" then
								if SN > "BRATHIGH" then
									if SN == "BRATLOW" then
										return CalcStat("BratProg",L,CalcStat("ProgBLow",L));
									else
										return 0;
									end
								elseif SN == "BRATHIGH" then
									return CalcStat("BratProg",L,CalcStat("ProgBHigh",L));
								else
									return 0;
								end
							else
								return CalcStat("BPEPRatPB",L)*CalcStat("BPEPRatPC",L);
							end
						else
							return 1;
						end
					elseif SN > "BRATMASTERY" then
						if SN < "BRATMITMEDIUM" then
							if SN < "BRATMITHEAVY" then
								if SN == "BRATMEDIUM" then
									return CalcStat("BratProg",L,CalcStat("ProgBMedium",L));
								else
									return 0;
								end
							elseif SN > "BRATMITHEAVY" then
								if SN == "BRATMITLIGHT" then
									return CalcStat("BratProg",L,CalcStat("ProgBMitLight",L));
								else
									return 0;
								end
							else
								return CalcStat("BratProg",L,CalcStat("ProgBMitHeavy",L));
							end
						elseif SN > "BRATMITMEDIUM" then
							if SN < "CRITDEFPBONUS" then
								if SN == "BRATPROG" then
									if L-DblCalcDev <= 75 then
										return LinFmod(RoundDbl(N),1,75,1,75,L);
									elseif L-DblCalcDev <= 76 then
										return LinFmod(1,RoundDbl(N)*75,CalcStat("StdProg",76,N),75,76,L);
									elseif L-DblCalcDev <= 100 then
										return LinFmod(1,CalcStat("StdProg",76,N),CalcStat("StdProg",100,N),75,100,L,10);
									elseif L-DblCalcDev <= 101 then
										return CalcStat("StdProg",L,N);
									elseif L-DblCalcDev <= 105 then
										return LinFmod(1,CalcStat("StdProg",101,N),CalcStat("StdProg",105,N),100,105,L,10);
									elseif L-DblCalcDev <= 106 then
										return CalcStat("StdProg",L,N);
									elseif L-DblCalcDev <= 115 then
										return LinFmod(1,CalcStat("StdProg",106,N),CalcStat("StdProg",115,N),106,115,L,10);
									elseif L-DblCalcDev <= 116 then
										return CalcStat("StdProg",L,N);
									elseif L-DblCalcDev <= 120 then
										return LinFmod(1,CalcStat("StdProg",116,N),CalcStat("StdProg",120,N),116,120,L,10);
									elseif L-DblCalcDev <= 121 then
										return CalcStat("StdProg",L,N);
									elseif L-DblCalcDev <= 130 then
										return LinFmod(1,CalcStat("StdProg",121,N),CalcStat("StdProg",130,N),121,130,L,10);
									elseif L-DblCalcDev <= 131 then
										return CalcStat("StdProg",L,N);
									else
										return LinFmod(1,CalcStat("StdProg",131,N),CalcStat("StdProg",140,N),131,140,L,10);
									end
								else
									return 0;
								end
							elseif SN > "CRITDEFPBONUS" then
								if SN > "CRITDEFPPRAT" then
									if SN == "CRITDEFPRATP" then
										return CalcPercAB(CalcStat("CritDefPRatPA",L),CalcStat("CritDefPRatPB",L),CalcStat("CritDefPRatPCap",L),N);
									else
										return 0;
									end
								elseif SN == "CRITDEFPPRAT" then
									return CalcRatAB(CalcStat("CritDefPRatPA",L),CalcStat("CritDefPRatPB",L),CalcStat("CritDefPRatPCapR",L),N);
								else
									return 0;
								end
							else
								return 0;
							end
						else
							return CalcStat("BratProg",L,CalcStat("ProgBMitMedium",L));
						end
					else
						return CalcStat("BratProg",L,CalcStat("ProgBMastery",L));
					end
				else
					return CalcRatAB(CalcStat("BPEPRatPA",L),CalcStat("BPEPRatPB",L),CalcStat("BPEPRatPCapR",L),N);
				end
			elseif SN > "CRITDEFPRATPA" then
				if SN < "CRITMAGNPRATPC" then
					if SN < "CRITHITPRATPB" then
						if SN < "CRITDEFPRATPCAPR" then
							if SN < "CRITDEFPRATPC" then
								if SN == "CRITDEFPRATPB" then
									return CalcStat("BratLow",L);
								else
									return 0;
								end
							elseif SN > "CRITDEFPRATPC" then
								if SN == "CRITDEFPRATPCAP" then
									return 80;
								else
									return 0;
								end
							else
								return 1;
							end
						elseif SN > "CRITDEFPRATPCAPR" then
							if SN < "CRITHITPPRAT" then
								if SN == "CRITHITPBONUS" then
									return 0;
								else
									return 0;
								end
							elseif SN > "CRITHITPPRAT" then
								if SN > "CRITHITPRATP" then
									if SN == "CRITHITPRATPA" then
										return 50;
									else
										return 0;
									end
								elseif SN == "CRITHITPRATP" then
									return CalcPercAB(CalcStat("CritHitPRatPA",L),CalcStat("CritHitPRatPB",L),CalcStat("CritHitPRatPCap",L),N);
								else
									return 0;
								end
							else
								return CalcRatAB(CalcStat("CritHitPRatPA",L),CalcStat("CritHitPRatPB",L),CalcStat("CritHitPRatPCapR",L),N);
							end
						else
							return CalcStat("CritDefPRatPB",L)*CalcStat("CritDefPRatPC",L);
						end
					elseif SN > "CRITHITPRATPB" then
						if SN < "CRITMAGNPBONUS" then
							if SN < "CRITHITPRATPCAP" then
								if SN == "CRITHITPRATPC" then
									return 1;
								else
									return 0;
								end
							elseif SN > "CRITHITPRATPCAP" then
								if SN == "CRITHITPRATPCAPR" then
									return CalcStat("CritHitPRatPB",L)*CalcStat("CritHitPRatPC",L);
								else
									return 0;
								end
							else
								return 25;
							end
						elseif SN > "CRITMAGNPBONUS" then
							if SN < "CRITMAGNPRATP" then
								if SN == "CRITMAGNPPRAT" then
									return CalcRatAB(CalcStat("CritMagnPRatPA",L),CalcStat("CritMagnPRatPB",L),CalcStat("CritMagnPRatPCapR",L),N);
								else
									return 0;
								end
							elseif SN > "CRITMAGNPRATP" then
								if SN > "CRITMAGNPRATPA" then
									if SN == "CRITMAGNPRATPB" then
										return CalcStat("BratHigh",L);
									else
										return 0;
									end
								elseif SN == "CRITMAGNPRATPA" then
									if L-DblCalcDev <= 120 then
										return 200;
									elseif L-DblCalcDev <= 127 then
										return (-5)*L+750;
									else
										return 112.5;
									end
								else
									return 0;
								end
							else
								return CalcPercAB(CalcStat("CritMagnPRatPA",L),CalcStat("CritMagnPRatPB",L),CalcStat("CritMagnPRatPCap",L),N);
							end
						else
							return 0;
						end
					else
						return CalcStat("BratLow",L);
					end
				elseif SN > "CRITMAGNPRATPC" then
					if SN < "DEVHITPRATPCAP" then
						if SN < "DEVHITPPRAT" then
							if SN < "CRITMAGNPRATPCAPR" then
								if SN == "CRITMAGNPRATPCAP" then
									if L-DblCalcDev <= 120 then
										return 100;
									else
										return 75;
									end
								else
									return 0;
								end
							elseif SN > "CRITMAGNPRATPCAPR" then
								if SN == "DEVHITPBONUS" then
									return 0;
								else
									return 0;
								end
							else
								return CalcStat("CritMagnPRatPB",L)*CalcStat("CritMagnPRatPC",L);
							end
						elseif SN > "DEVHITPPRAT" then
							if SN < "DEVHITPRATPA" then
								if SN == "DEVHITPRATP" then
									return CalcPercAB(CalcStat("DevHitPRatPA",L),CalcStat("DevHitPRatPB",L),CalcStat("DevHitPRatPCap",L),N);
								else
									return 0;
								end
							elseif SN > "DEVHITPRATPA" then
								if SN > "DEVHITPRATPB" then
									if SN == "DEVHITPRATPC" then
										return 1;
									else
										return 0;
									end
								elseif SN == "DEVHITPRATPB" then
									return CalcStat("BratMedium",L);
								else
									return 0;
								end
							else
								return 20;
							end
						else
							return CalcRatAB(CalcStat("DevHitPRatPA",L),CalcStat("DevHitPRatPB",L),CalcStat("DevHitPRatPCapR",L),N);
						end
					elseif SN > "DEVHITPRATPCAP" then
						if SN < "EVADEPRATP" then
							if SN < "EVADEPBONUS" then
								if SN == "DEVHITPRATPCAPR" then
									return CalcStat("DevHitPRatPB",L)*CalcStat("DevHitPRatPC",L);
								else
									return 0;
								end
							elseif SN > "EVADEPBONUS" then
								if SN == "EVADEPPRAT" then
									return CalcStat("BPEPPRat",L,N);
								else
									return 0;
								end
							else
								return CalcStat("BPEPBonus",L);
							end
						elseif SN > "EVADEPRATP" then
							if SN < "EVADEPRATPB" then
								if SN == "EVADEPRATPA" then
									return CalcStat("BPEPRatPA",L);
								else
									return 0;
								end
							elseif SN > "EVADEPRATPB" then
								if SN > "EVADEPRATPC" then
									if SN == "EVADEPRATPCAP" then
										return CalcStat("BPEPRatPCap",L);
									else
										return 0;
									end
								elseif SN == "EVADEPRATPC" then
									return CalcStat("BPEPRatPC",L);
								else
									return 0;
								end
							else
								return CalcStat("BPEPRatPB",L);
							end
						else
							return CalcStat("BPEPRatP",L,N);
						end
					else
						return 10;
					end
				else
					return CalcStat("CritMagnPRatPCap",L)/(CalcStat("CritMagnPRatPA",L)-CalcStat("CritMagnPRatPCap",L));
				end
			else
				return 160;
			end
		elseif SN > "EVADEPRATPCAPR" then
			if SN < "MITMEDIUMPPRAT" then
				if SN < "ITEMPROG" then
					if SN < "ILVLTOLVL" then
						if SN < "FINESSEPRATPA" then
							if SN < "FINESSEPPRAT" then
								if SN == "FINESSEPBONUS" then
									return 0;
								else
									return 0;
								end
							elseif SN > "FINESSEPPRAT" then
								if SN == "FINESSEPRATP" then
									return CalcPercAB(CalcStat("FinessePRatPA",L),CalcStat("FinessePRatPB",L),CalcStat("FinessePRatPCap",L),N);
								else
									return 0;
								end
							else
								return CalcRatAB(CalcStat("FinessePRatPA",L),CalcStat("FinessePRatPB",L),CalcStat("FinessePRatPCapR",L),N);
							end
						elseif SN > "FINESSEPRATPA" then
							if SN < "FINESSEPRATPC" then
								if SN == "FINESSEPRATPB" then
									return CalcStat("BratLow",L);
								else
									return 0;
								end
							elseif SN > "FINESSEPRATPC" then
								if SN > "FINESSEPRATPCAP" then
									if SN == "FINESSEPRATPCAPR" then
										return CalcStat("FinessePRatPB",L)*CalcStat("FinessePRatPC",L);
									else
										return 0;
									end
								elseif SN == "FINESSEPRATPCAP" then
									return 50;
								else
									return 0;
								end
							else
								return 1;
							end
						else
							return 100;
						end
					elseif SN > "ILVLTOLVL" then
						if SN < "INHEALPRATPA" then
							if SN < "INHEALPPRAT" then
								if SN == "INHEALPBONUS" then
									return 0;
								else
									return 0;
								end
							elseif SN > "INHEALPPRAT" then
								if SN == "INHEALPRATP" then
									return CalcPercAB(CalcStat("InHealPRatPA",L),CalcStat("InHealPRatPB",L),CalcStat("InHealPRatPCap",L),N);
								else
									return 0;
								end
							else
								return CalcRatAB(CalcStat("InHealPRatPA",L),CalcStat("InHealPRatPB",L),CalcStat("InHealPRatPCapR",L),N);
							end
						elseif SN > "INHEALPRATPA" then
							if SN < "INHEALPRATPC" then
								if SN == "INHEALPRATPB" then
									return CalcStat("BratLow",L);
								else
									return 0;
								end
							elseif SN > "INHEALPRATPC" then
								if SN > "INHEALPRATPCAP" then
									if SN == "INHEALPRATPCAPR" then
										return CalcStat("InHealPRatPB",L)*CalcStat("InHealPRatPC",L);
									else
										return 0;
									end
								elseif SN == "INHEALPRATPCAP" then
									return 25;
								else
									return 0;
								end
							else
								return 1;
							end
						else
							return 50;
						end
					else
						if L-DblCalcDev <= 79 then
							return LinFmod(1,1,75,1,79,L)*N;
						elseif L-DblCalcDev <= 80 then
							return LinFmod(1,75,76,79,80,L)*N;
						elseif L-DblCalcDev <= 200 then
							return LinFmod(1,76,100,80,200,L)*N;
						elseif L-DblCalcDev <= 205 then
							return LinFmod(1,100,101,200,205,L)*N;
						elseif L-DblCalcDev <= 225 then
							return LinFmod(1,101,105,205,225,L)*N;
						elseif L-DblCalcDev <= 300 then
							return LinFmod(1,105,106,225,300,L)*N;
						elseif L-DblCalcDev <= 349 then
							return LinFmod(1,106,115,300,349,L)*N;
						elseif L-DblCalcDev <= 350 then
							return LinFmod(1,115,116,349,350,L)*N;
						elseif L-DblCalcDev <= 399 then
							return LinFmod(1,116,120,350,399,L)*N;
						elseif L-DblCalcDev <= 400 then
							return LinFmod(1,120,121,399,400,L)*N;
						elseif L-DblCalcDev <= 449 then
							return LinFmod(1,121,130,400,449,L)*N;
						elseif L-DblCalcDev <= 450 then
							return LinFmod(1,130,131,449,450,L)*N;
						else
							return LinFmod(1,131,140,450,499,L)*N;
						end
					end
				elseif SN > "ITEMPROG" then
					if SN < "MITLIGHTPBONUS" then
						if SN < "MITHEAVYPRATPA" then
							if SN < "MITHEAVYPPRAT" then
								if SN == "MITHEAVYPBONUS" then
									return 0;
								else
									return 0;
								end
							elseif SN > "MITHEAVYPPRAT" then
								if SN == "MITHEAVYPRATP" then
									return CalcPercAB(CalcStat("MitHeavyPRatPA",L),CalcStat("MitHeavyPRatPB",L),CalcStat("MitHeavyPRatPCap",L),N);
								else
									return 0;
								end
							else
								return CalcRatAB(CalcStat("MitHeavyPRatPA",L),CalcStat("MitHeavyPRatPB",L),CalcStat("MitHeavyPRatPCapR",L),N);
							end
						elseif SN > "MITHEAVYPRATPA" then
							if SN < "MITHEAVYPRATPC" then
								if SN == "MITHEAVYPRATPB" then
									return CalcStat("BratMitHeavy",L);
								else
									return 0;
								end
							elseif SN > "MITHEAVYPRATPC" then
								if SN > "MITHEAVYPRATPCAP" then
									if SN == "MITHEAVYPRATPCAPR" then
										return CalcStat("MitHeavyPRatPB",L)*CalcStat("MitHeavyPRatPC",L);
									else
										return 0;
									end
								elseif SN == "MITHEAVYPRATPCAP" then
									return 60;
								else
									return 0;
								end
							else
								return 1.2;
							end
						else
							return 110;
						end
					elseif SN > "MITLIGHTPBONUS" then
						if SN < "MITLIGHTPRATPB" then
							if SN < "MITLIGHTPRATP" then
								if SN == "MITLIGHTPPRAT" then
									return CalcRatAB(CalcStat("MitLightPRatPA",L),CalcStat("MitLightPRatPB",L),CalcStat("MitLightPRatPCapR",L),N);
								else
									return 0;
								end
							elseif SN > "MITLIGHTPRATP" then
								if SN == "MITLIGHTPRATPA" then
									return 65;
								else
									return 0;
								end
							else
								return CalcPercAB(CalcStat("MitLightPRatPA",L),CalcStat("MitLightPRatPB",L),CalcStat("MitLightPRatPCap",L),N);
							end
						elseif SN > "MITLIGHTPRATPB" then
							if SN < "MITLIGHTPRATPCAP" then
								if SN == "MITLIGHTPRATPC" then
									return 1.6;
								else
									return 0;
								end
							elseif SN > "MITLIGHTPRATPCAP" then
								if SN > "MITLIGHTPRATPCAPR" then
									if SN == "MITMEDIUMPBONUS" then
										return 0;
									else
										return 0;
									end
								elseif SN == "MITLIGHTPRATPCAPR" then
									return CalcStat("MitLightPRatPB",L)*CalcStat("MitLightPRatPC",L);
								else
									return 0;
								end
							else
								return 40;
							end
						else
							return CalcStat("BratMitLight",L);
						end
					else
						return 0;
					end
				else
					return CalcStat("StatProg",CalcStat("ILvlToLvl",L),N);
				end
			elseif SN > "MITMEDIUMPPRAT" then
				if SN < "PARRYPRATPA" then
					if SN < "OUTHEALPRATP" then
						if SN < "MITMEDIUMPRATPC" then
							if SN < "MITMEDIUMPRATPA" then
								if SN == "MITMEDIUMPRATP" then
									return CalcPercAB(CalcStat("MitMediumPRatPA",L),CalcStat("MitMediumPRatPB",L),CalcStat("MitMediumPRatPCap",L),N);
								else
									return 0;
								end
							elseif SN > "MITMEDIUMPRATPA" then
								if SN == "MITMEDIUMPRATPB" then
									return CalcStat("BratMitMedium",L);
								else
									return 0;
								end
							else
								return 85;
							end
						elseif SN > "MITMEDIUMPRATPC" then
							if SN < "MITMEDIUMPRATPCAPR" then
								if SN == "MITMEDIUMPRATPCAP" then
									return 50;
								else
									return 0;
								end
							elseif SN > "MITMEDIUMPRATPCAPR" then
								if SN > "OUTHEALPBONUS" then
									if SN == "OUTHEALPPRAT" then
										return CalcRatAB(CalcStat("OutHealPRatPA",L),CalcStat("OutHealPRatPB",L),CalcStat("OutHealPRatPCapR",L),N);
									else
										return 0;
									end
								elseif SN == "OUTHEALPBONUS" then
									return 0;
								else
									return 0;
								end
							else
								return CalcStat("MitMediumPRatPB",L)*CalcStat("MitMediumPRatPC",L);
							end
						else
							return 10/7;
						end
					elseif SN > "OUTHEALPRATP" then
						if SN < "OUTHEALPRATPCAP" then
							if SN < "OUTHEALPRATPB" then
								if SN == "OUTHEALPRATPA" then
									return 140;
								else
									return 0;
								end
							elseif SN > "OUTHEALPRATPB" then
								if SN == "OUTHEALPRATPC" then
									return 1;
								else
									return 0;
								end
							else
								return CalcStat("BratMedium",L);
							end
						elseif SN > "OUTHEALPRATPCAP" then
							if SN < "PARRYPBONUS" then
								if SN == "OUTHEALPRATPCAPR" then
									return CalcStat("OutHealPRatPB",L)*CalcStat("OutHealPRatPC",L);
								else
									return 0;
								end
							elseif SN > "PARRYPBONUS" then
								if SN > "PARRYPPRAT" then
									if SN == "PARRYPRATP" then
										return CalcStat("BPEPRatP",L,N);
									else
										return 0;
									end
								elseif SN == "PARRYPPRAT" then
									return CalcStat("BPEPPRat",L,N);
								else
									return 0;
								end
							else
								return CalcStat("BPEPBonus",L);
							end
						else
							return 70;
						end
					else
						return CalcPercAB(CalcStat("OutHealPRatPA",L),CalcStat("OutHealPRatPB",L),CalcStat("OutHealPRatPCap",L),N);
					end
				elseif SN > "PARRYPRATPA" then
					if SN < "PARTBLOCKMITPRATPB" then
						if SN < "PARRYPRATPCAPR" then
							if SN < "PARRYPRATPC" then
								if SN == "PARRYPRATPB" then
									return CalcStat("BPEPRatPB",L);
								else
									return 0;
								end
							elseif SN > "PARRYPRATPC" then
								if SN == "PARRYPRATPCAP" then
									return CalcStat("BPEPRatPCap",L);
								else
									return 0;
								end
							else
								return CalcStat("BPEPRatPC",L);
							end
						elseif SN > "PARRYPRATPCAPR" then
							if SN < "PARTBLOCKMITPPRAT" then
								if SN == "PARTBLOCKMITPBONUS" then
									return CalcStat("PartMitPBonus",L);
								else
									return 0;
								end
							elseif SN > "PARTBLOCKMITPPRAT" then
								if SN > "PARTBLOCKMITPRATP" then
									if SN == "PARTBLOCKMITPRATPA" then
										return CalcStat("PartMitPRatPA",L);
									else
										return 0;
									end
								elseif SN == "PARTBLOCKMITPRATP" then
									return CalcStat("PartMitPRatP",L,N);
								else
									return 0;
								end
							else
								return CalcStat("PartMitPPRat",L,N);
							end
						else
							return CalcStat("BPEPRatPCapR",L);
						end
					elseif SN > "PARTBLOCKMITPRATPB" then
						if SN < "PARTBLOCKPBONUS" then
							if SN < "PARTBLOCKMITPRATPCAP" then
								if SN == "PARTBLOCKMITPRATPC" then
									return CalcStat("PartMitPRatPC",L);
								else
									return 0;
								end
							elseif SN > "PARTBLOCKMITPRATPCAP" then
								if SN == "PARTBLOCKMITPRATPCAPR" then
									return CalcStat("PartMitPRatPCapR",L);
								else
									return 0;
								end
							else
								return CalcStat("PartMitPRatPCap",L);
							end
						elseif SN > "PARTBLOCKPBONUS" then
							if SN < "PARTBLOCKPRATP" then
								if SN == "PARTBLOCKPPRAT" then
									return CalcStat("PartBPEPPRat",L,N);
								else
									return 0;
								end
							elseif SN > "PARTBLOCKPRATP" then
								if SN > "PARTBLOCKPRATPA" then
									if SN == "PARTBLOCKPRATPB" then
										return CalcStat("PartBPEPRatPB",L);
									else
										return 0;
									end
								elseif SN == "PARTBLOCKPRATPA" then
									return CalcStat("PartBPEPRatPA",L);
								else
									return 0;
								end
							else
								return CalcStat("PartBPEPRatP",L,N);
							end
						else
							return CalcStat("PartBPEPBonus",L);
						end
					else
						return CalcStat("PartMitPRatPB",L);
					end
				else
					return CalcStat("BPEPRatPA",L);
				end
			else
				return CalcRatAB(CalcStat("MitMediumPRatPA",L),CalcStat("MitMediumPRatPB",L),CalcStat("MitMediumPRatPCapR",L),N);
			end
		else
			return CalcStat("BPEPRatPCapR",L);
		end
	elseif SN > "PARTBLOCKPRATPC" then
		if SN < "PHYMITLPRATPC" then
			if SN < "PARTPARRYMITPPRAT" then
				if SN < "PARTEVADEMITPRATPCAPR" then
					if SN < "PARTBPEPRATPCAP" then
						if SN < "PARTBPEPPRAT" then
							if SN < "PARTBLOCKPRATPCAPR" then
								if SN == "PARTBLOCKPRATPCAP" then
									return CalcStat("PartBPEPRatPCap",L);
								else
									return 0;
								end
							elseif SN > "PARTBLOCKPRATPCAPR" then
								if SN == "PARTBPEPBONUS" then
									return 0;
								else
									return 0;
								end
							else
								return CalcStat("PartBPEPRatPCapR",L);
							end
						elseif SN > "PARTBPEPPRAT" then
							if SN < "PARTBPEPRATPA" then
								if SN == "PARTBPEPRATP" then
									return CalcPercAB(CalcStat("PartBPEPRatPA",L),CalcStat("PartBPEPRatPB",L),CalcStat("PartBPEPRatPCap",L),N);
								else
									return 0;
								end
							elseif SN > "PARTBPEPRATPA" then
								if SN > "PARTBPEPRATPB" then
									if SN == "PARTBPEPRATPC" then
										return 1;
									else
										return 0;
									end
								elseif SN == "PARTBPEPRATPB" then
									return CalcStat("BratMedium",L);
								else
									return 0;
								end
							else
								return 70;
							end
						else
							return CalcRatAB(CalcStat("PartBPEPRatPA",L),CalcStat("PartBPEPRatPB",L),CalcStat("PartBPEPRatPCapR",L),N);
						end
					elseif SN > "PARTBPEPRATPCAP" then
						if SN < "PARTEVADEMITPRATP" then
							if SN < "PARTEVADEMITPBONUS" then
								if SN == "PARTBPEPRATPCAPR" then
									return CalcStat("PartBPEPRatPB",L)*CalcStat("PartBPEPRatPC",L);
								else
									return 0;
								end
							elseif SN > "PARTEVADEMITPBONUS" then
								if SN == "PARTEVADEMITPPRAT" then
									return CalcStat("PartMitPPRat",L,N);
								else
									return 0;
								end
							else
								if L-DblCalcDev <= 1 then
									return 35;
								else
									return CalcStat("PartMitPBonus",L);
								end
							end
						elseif SN > "PARTEVADEMITPRATP" then
							if SN < "PARTEVADEMITPRATPB" then
								if SN == "PARTEVADEMITPRATPA" then
									return CalcStat("PartMitPRatPA",L);
								else
									return 0;
								end
							elseif SN > "PARTEVADEMITPRATPB" then
								if SN > "PARTEVADEMITPRATPC" then
									if SN == "PARTEVADEMITPRATPCAP" then
										return CalcStat("PartMitPRatPCap",L);
									else
										return 0;
									end
								elseif SN == "PARTEVADEMITPRATPC" then
									return CalcStat("PartMitPRatPC",L);
								else
									return 0;
								end
							else
								return CalcStat("PartMitPRatPB",L);
							end
						else
							return CalcStat("PartMitPRatP",L,N);
						end
					else
						return 35;
					end
				elseif SN > "PARTEVADEMITPRATPCAPR" then
					if SN < "PARTMITPBONUS" then
						if SN < "PARTEVADEPRATPA" then
							if SN < "PARTEVADEPPRAT" then
								if SN == "PARTEVADEPBONUS" then
									return CalcStat("PartBPEPBonus",L);
								else
									return 0;
								end
							elseif SN > "PARTEVADEPPRAT" then
								if SN == "PARTEVADEPRATP" then
									return CalcStat("PartBPEPRatP",L,N);
								else
									return 0;
								end
							else
								return CalcStat("PartBPEPPRat",L,N);
							end
						elseif SN > "PARTEVADEPRATPA" then
							if SN < "PARTEVADEPRATPC" then
								if SN == "PARTEVADEPRATPB" then
									return CalcStat("PartBPEPRatPB",L);
								else
									return 0;
								end
							elseif SN > "PARTEVADEPRATPC" then
								if SN > "PARTEVADEPRATPCAP" then
									if SN == "PARTEVADEPRATPCAPR" then
										return CalcStat("PartBPEPRatPCapR",L);
									else
										return 0;
									end
								elseif SN == "PARTEVADEPRATPCAP" then
									return CalcStat("PartBPEPRatPCap",L);
								else
									return 0;
								end
							else
								return CalcStat("PartBPEPRatPC",L);
							end
						else
							return CalcStat("PartBPEPRatPA",L);
						end
					elseif SN > "PARTMITPBONUS" then
						if SN < "PARTMITPRATPB" then
							if SN < "PARTMITPRATP" then
								if SN == "PARTMITPPRAT" then
									return CalcRatAB(CalcStat("PartMitPRatPA",L),CalcStat("PartMitPRatPB",L),CalcStat("PartMitPRatPCapR",L),N);
								else
									return 0;
								end
							elseif SN > "PARTMITPRATP" then
								if SN == "PARTMITPRATPA" then
									return 100;
								else
									return 0;
								end
							else
								return CalcPercAB(CalcStat("PartMitPRatPA",L),CalcStat("PartMitPRatPB",L),CalcStat("PartMitPRatPCap",L),N);
							end
						elseif SN > "PARTMITPRATPB" then
							if SN < "PARTMITPRATPCAP" then
								if SN == "PARTMITPRATPC" then
									return 1;
								else
									return 0;
								end
							elseif SN > "PARTMITPRATPCAP" then
								if SN > "PARTMITPRATPCAPR" then
									if SN == "PARTPARRYMITPBONUS" then
										return CalcStat("PartMitPBonus",L);
									else
										return 0;
									end
								elseif SN == "PARTMITPRATPCAPR" then
									return CalcStat("PartMitPRatPB",L)*CalcStat("PartMitPRatPC",L);
								else
									return 0;
								end
							else
								return 50;
							end
						else
							return CalcStat("BratMedium",L);
						end
					else
						return 10;
					end
				else
					return CalcStat("PartMitPRatPCapR",L);
				end
			elseif SN > "PARTPARRYMITPPRAT" then
				if SN < "PHYDMGPRATPA" then
					if SN < "PARTPARRYPRATP" then
						if SN < "PARTPARRYMITPRATPC" then
							if SN < "PARTPARRYMITPRATPA" then
								if SN == "PARTPARRYMITPRATP" then
									return CalcStat("PartMitPRatP",L,N);
								else
									return 0;
								end
							elseif SN > "PARTPARRYMITPRATPA" then
								if SN == "PARTPARRYMITPRATPB" then
									return CalcStat("PartMitPRatPB",L);
								else
									return 0;
								end
							else
								return CalcStat("PartMitPRatPA",L);
							end
						elseif SN > "PARTPARRYMITPRATPC" then
							if SN < "PARTPARRYMITPRATPCAPR" then
								if SN == "PARTPARRYMITPRATPCAP" then
									return CalcStat("PartMitPRatPCap",L);
								else
									return 0;
								end
							elseif SN > "PARTPARRYMITPRATPCAPR" then
								if SN > "PARTPARRYPBONUS" then
									if SN == "PARTPARRYPPRAT" then
										return CalcStat("PartBPEPPRat",L,N);
									else
										return 0;
									end
								elseif SN == "PARTPARRYPBONUS" then
									return CalcStat("PartBPEPBonus",L);
								else
									return 0;
								end
							else
								return CalcStat("PartMitPRatPCapR",L);
							end
						else
							return CalcStat("PartMitPRatPC",L);
						end
					elseif SN > "PARTPARRYPRATP" then
						if SN < "PARTPARRYPRATPCAP" then
							if SN < "PARTPARRYPRATPB" then
								if SN == "PARTPARRYPRATPA" then
									return CalcStat("PartBPEPRatPA",L);
								else
									return 0;
								end
							elseif SN > "PARTPARRYPRATPB" then
								if SN == "PARTPARRYPRATPC" then
									return CalcStat("PartBPEPRatPC",L);
								else
									return 0;
								end
							else
								return CalcStat("PartBPEPRatPB",L);
							end
						elseif SN > "PARTPARRYPRATPCAP" then
							if SN < "PHYDMGPBONUS" then
								if SN == "PARTPARRYPRATPCAPR" then
									return CalcStat("PartBPEPRatPCapR",L);
								else
									return 0;
								end
							elseif SN > "PHYDMGPBONUS" then
								if SN > "PHYDMGPPRAT" then
									if SN == "PHYDMGPRATP" then
										return CalcPercAB(CalcStat("PhyDmgPRatPA",L),CalcStat("PhyDmgPRatPB",L),CalcStat("PhyDmgPRatPCap",L),N);
									else
										return 0;
									end
								elseif SN == "PHYDMGPPRAT" then
									return CalcRatAB(CalcStat("PhyDmgPRatPA",L),CalcStat("PhyDmgPRatPB",L),CalcStat("PhyDmgPRatPCapR",L),N);
								else
									return 0;
								end
							else
								return 0;
							end
						else
							return CalcStat("PartBPEPRatPCap",L);
						end
					else
						return CalcStat("PartBPEPRatP",L,N);
					end
				elseif SN > "PHYDMGPRATPA" then
					if SN < "PHYMITHPRATPB" then
						if SN < "PHYDMGPRATPCAPR" then
							if SN < "PHYDMGPRATPC" then
								if SN == "PHYDMGPRATPB" then
									return CalcStat("BratMastery",L);
								else
									return 0;
								end
							elseif SN > "PHYDMGPRATPC" then
								if SN == "PHYDMGPRATPCAP" then
									return 200;
								else
									return 0;
								end
							else
								return 1;
							end
						elseif SN > "PHYDMGPRATPCAPR" then
							if SN < "PHYMITHPPRAT" then
								if SN == "PHYMITHPBONUS" then
									return CalcStat("MitHeavyPBonus",L);
								else
									return 0;
								end
							elseif SN > "PHYMITHPPRAT" then
								if SN > "PHYMITHPRATP" then
									if SN == "PHYMITHPRATPA" then
										return CalcStat("MitHeavyPRatPA",L);
									else
										return 0;
									end
								elseif SN == "PHYMITHPRATP" then
									return CalcStat("MitHeavyPRatP",L,N);
								else
									return 0;
								end
							else
								return CalcStat("MitHeavyPPRat",L,N);
							end
						else
							return CalcStat("PhyDmgPRatPB",L)*CalcStat("PhyDmgPRatPC",L);
						end
					elseif SN > "PHYMITHPRATPB" then
						if SN < "PHYMITLPBONUS" then
							if SN < "PHYMITHPRATPCAP" then
								if SN == "PHYMITHPRATPC" then
									return CalcStat("MitHeavyPRatPC",L);
								else
									return 0;
								end
							elseif SN > "PHYMITHPRATPCAP" then
								if SN == "PHYMITHPRATPCAPR" then
									return CalcStat("MitHeavyPRatPCapR",L);
								else
									return 0;
								end
							else
								return CalcStat("MitHeavyPRatPCap",L);
							end
						elseif SN > "PHYMITLPBONUS" then
							if SN < "PHYMITLPRATP" then
								if SN == "PHYMITLPPRAT" then
									return CalcStat("MitLightPPRat",L,N);
								else
									return 0;
								end
							elseif SN > "PHYMITLPRATP" then
								if SN > "PHYMITLPRATPA" then
									if SN == "PHYMITLPRATPB" then
										return CalcStat("MitLightPRatPB",L);
									else
										return 0;
									end
								elseif SN == "PHYMITLPRATPA" then
									return CalcStat("MitLightPRatPA",L);
								else
									return 0;
								end
							else
								return CalcStat("MitLightPRatP",L,N);
							end
						else
							return CalcStat("MitLightPBonus",L);
						end
					else
						return CalcStat("MitHeavyPRatPB",L);
					end
				else
					return 400;
				end
			else
				return CalcStat("PartMitPPRat",L,N);
			end
		elseif SN > "PHYMITLPRATPC" then
			if SN < "TACDMGPPRAT" then
				if SN < "PROGBMITMEDIUM" then
					if SN < "PHYMITMPRATPCAP" then
						if SN < "PHYMITMPPRAT" then
							if SN < "PHYMITLPRATPCAPR" then
								if SN == "PHYMITLPRATPCAP" then
									return CalcStat("MitLightPRatPCap",L);
								else
									return 0;
								end
							elseif SN > "PHYMITLPRATPCAPR" then
								if SN == "PHYMITMPBONUS" then
									return CalcStat("MitMediumPBonus",L);
								else
									return 0;
								end
							else
								return CalcStat("MitLightPRatPCapR",L);
							end
						elseif SN > "PHYMITMPPRAT" then
							if SN < "PHYMITMPRATPA" then
								if SN == "PHYMITMPRATP" then
									return CalcStat("MitMediumPRatP",L,N);
								else
									return 0;
								end
							elseif SN > "PHYMITMPRATPA" then
								if SN > "PHYMITMPRATPB" then
									if SN == "PHYMITMPRATPC" then
										return CalcStat("MitMediumPRatPC",L);
									else
										return 0;
									end
								elseif SN == "PHYMITMPRATPB" then
									return CalcStat("MitMediumPRatPB",L);
								else
									return 0;
								end
							else
								return CalcStat("MitMediumPRatPA",L);
							end
						else
							return CalcStat("MitMediumPPRat",L,N);
						end
					elseif SN > "PHYMITMPRATPCAP" then
						if SN < "PROGBLOW" then
							if SN < "PNTMPDEFENCE" then
								if SN == "PHYMITMPRATPCAPR" then
									return CalcStat("MitMediumPRatPCapR",L);
								else
									return 0;
								end
							elseif SN > "PNTMPDEFENCE" then
								if SN == "PROGBHIGH" then
									return 500;
								else
									return 0;
								end
							else
								return 351/13000;
							end
						elseif SN > "PROGBLOW" then
							if SN < "PROGBMEDIUM" then
								if SN == "PROGBMASTERY" then
									return 270;
								else
									return 0;
								end
							elseif SN > "PROGBMEDIUM" then
								if SN > "PROGBMITHEAVY" then
									if SN == "PROGBMITLIGHT" then
										return 280/3;
									else
										return 0;
									end
								elseif SN == "PROGBMITHEAVY" then
									return 174;
								else
									return 0;
								end
							else
								return 400;
							end
						else
							return 200;
						end
					else
						return CalcStat("MitMediumPRatPCap",L);
					end
				elseif SN > "PROGBMITMEDIUM" then
					if SN < "RESISTPRATPCAPR" then
						if SN < "RESISTPRATP" then
							if SN < "RESISTPBONUS" then
								if SN == "RATDEFENCET" then
									return CalcStat("PntMPDefence",L)*CalcStat("TraitProg",L,CalcStat("ProgBMedium",L))*CalcStat("AdjTraitProgRatings",L,CalcStat("ProgBMedium",L))*N;
								else
									return 0;
								end
							elseif SN > "RESISTPBONUS" then
								if SN == "RESISTPPRAT" then
									return CalcRatAB(CalcStat("ResistPRatPA",L),CalcStat("ResistPRatPB",L),CalcStat("ResistPRatPCapR",L),N);
								else
									return 0;
								end
							else
								return 0;
							end
						elseif SN > "RESISTPRATP" then
							if SN < "RESISTPRATPB" then
								if SN == "RESISTPRATPA" then
									return 100;
								else
									return 0;
								end
							elseif SN > "RESISTPRATPB" then
								if SN > "RESISTPRATPC" then
									if SN == "RESISTPRATPCAP" then
										return 50;
									else
										return 0;
									end
								elseif SN == "RESISTPRATPC" then
									return 1;
								else
									return 0;
								end
							else
								return CalcStat("BratLow",L);
							end
						else
							return CalcPercAB(CalcStat("ResistPRatPA",L),CalcStat("ResistPRatPB",L),CalcStat("ResistPRatPCap",L),N);
						end
					elseif SN > "RESISTPRATPCAPR" then
						if SN < "T2PENARMOUR" then
							if SN < "STATPROG" then
								if SN == "RESISTT" then
									return CalcStat("RatDefenceT",L,N);
								else
									return 0;
								end
							elseif SN > "STATPROG" then
								if SN == "STDPROG" then
									if L-DblCalcDev <= 75 then
										return LinFmod(N,1,75,1,75,L);
									elseif L-DblCalcDev <= 76 then
										return LinFmod(1,N*75,RoundDbl(N*82.5,-2),75,76,L);
									elseif L-DblCalcDev <= 100 then
										return LinFmod(1,RoundDbl(N*82.5,-2),N*150,76,100,L);
									elseif L-DblCalcDev <= 101 then
										return LinFmod(N,150,165,100,101,L);
									elseif L-DblCalcDev <= 105 then
										return LinFmod(N,165,225,101,105,L);
									elseif L-DblCalcDev <= 106 then
										return LinFmod(N,225,270,105,106,L);
									elseif L-DblCalcDev <= 115 then
										return LinFmod(N,270,450,106,115,L);
									elseif L-DblCalcDev <= 116 then
										return LinFmod(N,450,495,115,116,L);
									elseif L-DblCalcDev <= 120 then
										return LinFmod(N,495,1125,116,120,L);
									elseif L-DblCalcDev <= 121 then
										return LinFmod(N,1125,1575,120,121,L);
									elseif L-DblCalcDev <= 130 then
										return LinFmod(N,1575,3150,121,130,L);
									elseif L-DblCalcDev <= 131 then
										return LinFmod(N,3150,3780,130,131,L);
									else
										return LinFmod(N,3780,6300,130,140,L);
									end
								else
									return 0;
								end
							else
								if L-DblCalcDev <= 75 then
									return LinFmod(RoundDbl(N),1,75,1,75,L);
								elseif L-DblCalcDev <= 76 then
									return LinFmod(1,RoundDbl(N)*75,CalcStat("StdProg",76,N),75,76,L);
								else
									return CalcStat("StdProg",L,N);
								end
							end
						elseif SN > "T2PENARMOUR" then
							if SN < "T2PENMIT" then
								if SN == "T2PENBPE" then
									if L-DblCalcDev <= 115 then
										return (-40)*L;
									elseif L-DblCalcDev <= 116 then
										return ExpFmod(CalcStat("T2penBPE",115),116,20,L);
									elseif L-DblCalcDev <= 120 then
										return ExpFmod(CalcStat("T2penBPE",116),117,5.5,L);
									elseif L-DblCalcDev <= 121 then
										return ExpFmod(CalcStat("T2penBPE",120),121,20,L);
									elseif L-DblCalcDev <= 125 then
										return ExpFmod(CalcStat("T2penBPE",121),122,5.5,L);
									elseif L-DblCalcDev <= 126 then
										return ExpFmod(CalcStat("T2penBPE",125),126,20,L);
									else
										return ExpFmod(CalcStat("T2penBPE",126),127,5.5,L);
									end
								else
									return 0;
								end
							elseif SN > "T2PENMIT" then
								if SN > "T2PENRESIST" then
									if SN == "TACDMGPBONUS" then
										return 0;
									else
										return 0;
									end
								elseif SN == "T2PENRESIST" then
									if L-DblCalcDev <= 115 then
										return (-90)*L;
									elseif L-DblCalcDev <= 116 then
										return ExpFmod(CalcStat("T2penResist",115),116,20,L);
									elseif L-DblCalcDev <= 120 then
										return ExpFmod(CalcStat("T2penResist",116),117,5.5,L);
									elseif L-DblCalcDev <= 121 then
										return ExpFmod(CalcStat("T2penResist",120),121,20,L);
									elseif L-DblCalcDev <= 125 then
										return ExpFmod(CalcStat("T2penResist",121),122,5.5,L);
									elseif L-DblCalcDev <= 126 then
										return ExpFmod(CalcStat("T2penResist",125),126,20,L);
									else
										return ExpFmod(CalcStat("T2penResist",126),127,5.5,L);
									end
								else
									return 0;
								end
							else
								if L-DblCalcDev <= 115 then
									return FloorDbl(L*13.5)*-5;
								elseif L-DblCalcDev <= 116 then
									return ExpFmod(CalcStat("T2penMit",115),116,20,L);
								elseif L-DblCalcDev <= 120 then
									return ExpFmod(CalcStat("T2penMit",116),117,5.5,L);
								elseif L-DblCalcDev <= 121 then
									return ExpFmod(CalcStat("T2penMit",120),121,20,L);
								elseif L-DblCalcDev <= 125 then
									return ExpFmod(CalcStat("T2penMit",121),122,5.5,L);
								elseif L-DblCalcDev <= 126 then
									return ExpFmod(CalcStat("T2penMit",125),126,20,L);
								else
									return ExpFmod(CalcStat("T2penMit",126),127,5.5,L);
								end
							end
						else
							return CalcStat("T2penMit",L);
						end
					else
						return CalcStat("ResistPRatPB",L)*CalcStat("ResistPRatPC",L);
					end
				else
					return 382/3;
				end
			elseif SN > "TACDMGPPRAT" then
				if SN < "TACMITLPRATPA" then
					if SN < "TACMITHPRATP" then
						if SN < "TACDMGPRATPC" then
							if SN < "TACDMGPRATPA" then
								if SN == "TACDMGPRATP" then
									return CalcPercAB(CalcStat("TacDmgPRatPA",L),CalcStat("TacDmgPRatPB",L),CalcStat("TacDmgPRatPCap",L),N);
								else
									return 0;
								end
							elseif SN > "TACDMGPRATPA" then
								if SN == "TACDMGPRATPB" then
									return CalcStat("BratMastery",L);
								else
									return 0;
								end
							else
								return 400;
							end
						elseif SN > "TACDMGPRATPC" then
							if SN < "TACDMGPRATPCAPR" then
								if SN == "TACDMGPRATPCAP" then
									return 200;
								else
									return 0;
								end
							elseif SN > "TACDMGPRATPCAPR" then
								if SN > "TACMITHPBONUS" then
									if SN == "TACMITHPPRAT" then
										return CalcStat("MitHeavyPPRat",L,N);
									else
										return 0;
									end
								elseif SN == "TACMITHPBONUS" then
									return CalcStat("MitHeavyPBonus",L);
								else
									return 0;
								end
							else
								return CalcStat("TacDmgPRatPB",L)*CalcStat("TacDmgPRatPC",L);
							end
						else
							return 1;
						end
					elseif SN > "TACMITHPRATP" then
						if SN < "TACMITHPRATPCAP" then
							if SN < "TACMITHPRATPB" then
								if SN == "TACMITHPRATPA" then
									return CalcStat("MitHeavyPRatPA",L);
								else
									return 0;
								end
							elseif SN > "TACMITHPRATPB" then
								if SN == "TACMITHPRATPC" then
									return CalcStat("MitHeavyPRatPC",L);
								else
									return 0;
								end
							else
								return CalcStat("MitHeavyPRatPB",L);
							end
						elseif SN > "TACMITHPRATPCAP" then
							if SN < "TACMITLPBONUS" then
								if SN == "TACMITHPRATPCAPR" then
									return CalcStat("MitHeavyPRatPCapR",L);
								else
									return 0;
								end
							elseif SN > "TACMITLPBONUS" then
								if SN > "TACMITLPPRAT" then
									if SN == "TACMITLPRATP" then
										return CalcStat("MitLightPRatP",L,N);
									else
										return 0;
									end
								elseif SN == "TACMITLPPRAT" then
									return CalcStat("MitLightPPRat",L,N);
								else
									return 0;
								end
							else
								return CalcStat("MitLightPBonus",L);
							end
						else
							return CalcStat("MitHeavyPRatPCap",L);
						end
					else
						return CalcStat("MitHeavyPRatP",L,N);
					end
				elseif SN > "TACMITLPRATPA" then
					if SN < "TACMITMPRATPB" then
						if SN < "TACMITLPRATPCAPR" then
							if SN < "TACMITLPRATPC" then
								if SN == "TACMITLPRATPB" then
									return CalcStat("MitLightPRatPB",L);
								else
									return 0;
								end
							elseif SN > "TACMITLPRATPC" then
								if SN == "TACMITLPRATPCAP" then
									return CalcStat("MitLightPRatPCap",L);
								else
									return 0;
								end
							else
								return CalcStat("MitLightPRatPC",L);
							end
						elseif SN > "TACMITLPRATPCAPR" then
							if SN < "TACMITMPPRAT" then
								if SN == "TACMITMPBONUS" then
									return CalcStat("MitMediumPBonus",L);
								else
									return 0;
								end
							elseif SN > "TACMITMPPRAT" then
								if SN > "TACMITMPRATP" then
									if SN == "TACMITMPRATPA" then
										return CalcStat("MitMediumPRatPA",L);
									else
										return 0;
									end
								elseif SN == "TACMITMPRATP" then
									return CalcStat("MitMediumPRatP",L,N);
								else
									return 0;
								end
							else
								return CalcStat("MitMediumPPRat",L,N);
							end
						else
							return CalcStat("MitLightPRatPCapR",L);
						end
					elseif SN > "TACMITMPRATPB" then
						if SN < "TPENARMOUR" then
							if SN < "TACMITMPRATPCAP" then
								if SN == "TACMITMPRATPC" then
									return CalcStat("MitMediumPRatPC",L);
								else
									return 0;
								end
							elseif SN > "TACMITMPRATPCAP" then
								if SN == "TACMITMPRATPCAPR" then
									return CalcStat("MitMediumPRatPCapR",L);
								else
									return 0;
								end
							else
								return CalcStat("MitMediumPRatPCap",L);
							end
						elseif SN > "TPENARMOUR" then
							if SN < "TPENCHOICE" then
								if SN == "TPENBPE" then
									return CalcStat("TpenChoice",N)*CalcStat("RatDefenceT",L);
								else
									return 0;
								end
							elseif SN > "TPENCHOICE" then
								if SN > "TPENRESIST" then
									if SN == "TRAITPROG" then
										if L-DblCalcDev <= 105 then
											return LinFmod(1,CalcStat("StatProg",1,N),CalcStat("StatProg",105,N),1,105,L);
										else
											return CalcStat("StatProg",L,N);
										end
									else
										return 0;
									end
								elseif SN == "TPENRESIST" then
									return CalcStat("TpenChoice",N)*CalcStat("ResistT",L,2);
								else
									return 0;
								end
							else
								return DataTableValue({0,-1,-2},L);
							end
						else
							return CalcStat("TpenChoice",N)*(CalcStat("ArmourLowRaw",1,"MSh")/CalcStat("TraitProg",1,CalcStat("ProgBMitMedium",L)))*CalcStat("TraitProg",L,CalcStat("ProgBMitMedium",L))*CalcStat("AdjTraitProgRatings",L,CalcStat("ProgBMitMedium",L));
						end
					else
						return CalcStat("MitMediumPRatPB",L);
					end
				else
					return CalcStat("MitLightPRatPA",L);
				end
			else
				return CalcRatAB(CalcStat("TacDmgPRatPA",L),CalcStat("TacDmgPRatPB",L),CalcStat("TacDmgPRatPCapR",L),N);
			end
		else
			return CalcStat("MitLightPRatPC",L);
		end
	else
		return CalcStat("PartBPEPRatPC",L);
	end

end

-- Support functions for CalcStat. These consist of implementations of more complex calculation types, decode methods for parameter "C" and rounding/min/max/compare functions for floating point numbers.
-- Created by Giseldah

-- Floating point numbers bring errors into the calculation, both inside the Lotro-client and in this function collection. This is why a 100% match with the stats in Lotro is impossible.
-- Anyway, to compensate for some errors, we use a calculation deviation correction value. This makes for instance 24.49999999 round to 25, as it's assumed that 24.5 was intended as outcome of a formula.
DblCalcDev = 0.00000001;

-- ****************** Calculation Type support functions ******************

-- DDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDDD
-- DataTableValue: Takes a value from an array table.

function DataTableValue(vDataArray, Index)

	lI = Index;
	if lI < 1 then
		lI = 1;
	end
	if lI > #vDataArray then
		lI = #vDataArray;
	end

	return vDataArray[lI];

end

-- EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
-- ExpFmod: Exponential function based on percentage.
-- Common percentage values are around ~5.5% for between levels and ~20% jumps between level segments.

function ExpFmod(dVal, dLstart, dPlvl, dLvl)

	if SmallerDbl(dLvl,dLstart) then
		return dVal;
	else
		return dVal*(1+dPlvl/100)^(dLvl-dLstart+1);
	end

end

-- NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN
-- NamedRangeValue: Takes a value from a named spreadsheet table.
-- This function doesn't have a meaning outside a spreadsheet and so is not implemented here.

function NamedRangeValue(RName, RowIndex, ColIndex)

	return 0;

end

-- PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
-- CalcPercAB: Calculates the percentage out of a rating based on the AB formula.

function CalcPercAB(dA, dB, dPCap, dR)

	if dR <= 0 then
		return 0;
	else
		return MinDbl(dA/(1+dB/dR),dPCap);
	end

end

-- RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
-- CalcRatAB: Calculates the rating out of a percentage based on the AB formula.

function CalcRatAB(dA, dB, dCapR, dP)

	if dP <= 0 then
		return 0;
	else
		return MinDbl(dB/(dA/dP-1),dCapR);
	end
end

-- TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
-- LinFmod: Linear line function between 2 points with some optional modifications.
-- Connects point (dLstart,dVal*dFstart) with (dLend,dVal*dFend).
-- Usually used with dVal=1 and dFstart/dFend containing unrelated points or dVal=# and dFstart/dFend containing multiplier factors.
-- Modification for in-between points on the line: rounding (in general to multiples of dRound=10).

function LinFmod(dVal, dFstart, dFend, dLstart, dLend, dLvl, dRound)

	if IsDbl(dLvl,dLstart) then
		return dVal*dFstart;
	elseif IsDbl(dLvl,dLend) then
		return dVal*dFend;
	elseif dRound == nil then
		return dVal*(dFstart*(dLend-dLvl)+(dLvl-dLstart)*dFend)/(dLend-dLstart);
	else
		return RoundDbl((dVal*(dFstart*(dLend-dLvl)+(dLvl-dLstart)*dFend)/((dLend-dLstart)*dRound)))*dRound;
	end

end

-- ****************** Parameter "C" decode support functions ******************

-- ArmCodeIndex: returns a specified index from an Armour Code.
-- acode string:
-- 1st position: H=heavy, M=medium, L=light
-- 2nd position: H=head, S=shoulders, CL=cloak/back, C=chest, G=gloves, L=leggings, B=boots, Sh=shield
-- 3rd position: W=white/common, Y=yellow/uncommon, P=purple/rare, T=teal/blue/incomparable, G=gold/legendary/epic
-- Note: no such thing exists as a heavy, medium or light cloak, so no H/M/L in cloak codes (cloaks go automatically in the M class since U23, although historically this was L)

function ArmCodeIndex(acode, ii)

	local armourcode = string.upper(string.match(acode,"(%a+)"));

	-- get positional codes and make some corrections
	local sarmclass = string.sub(armourcode,1,1);
	local sarmtype = string.sub(armourcode,2,2);
	local sarmcol = string.sub(armourcode,3,3);
	if sarmtype == "S" and sarmcol == "H" then
		sarmtype = "SH";
		sarmcol = string.sub(armourcode,4,4);
	elseif sarmclass == "C" and sarmtype == "L" then
		sarmclass = "M";
		sarmtype = "CL";
	else
		sarmtype = " "..sarmtype;
	end
	
	if ii == 1 then
		return string.find("HML",sarmclass);
	elseif ii == 2 then
		return (string.find(" H SCL C G L BSH",sarmtype)+1)/2;
	elseif ii == 3 then
		return string.find("WYPTG",sarmcol);
	end
	
	return 0;
	
end

-- RomanRankDecode: converts a string with a Roman number in characters, to an integer number.
-- used for Legendary Item Title calculation.

local RomanCharsToValues = {["M"]=1000,["CM"]=900,["D"]=500,["CD"]=400,["C"]=100,["XC"]=90,["L"]=50,["XL"]=40,["X"]=10,["IX"]=9,["V"]=5,["IV"]=4,["I"]=1};

function RomanRankDecode(srank)

	if srank ~= nil then
		if srank ~= "" then
			for schars, ivalue in pairs(RomanCharsToValues) do
				if string.sub(string.upper(srank),1,string.len(schars)) == schars then
					return ivalue+RomanRankDecode(string.sub(srank,string.len(schars)+1));
				end
			end
		end
	end
	
	return 0;
	
end

-- ****************** Misc. floating point support functions ******************

-- Misc. functions for floating point: rounding etc.
-- For roundings: iDec is number of decimals.

function RoundDbl(dNum, iDec)

	if iDec == nil then
		iDec = 0;
	end
	if iDec == 0 then
		return math.floor(dNum+0.5+DblCalcDev);
	else
		return math.floor(dNum*10^iDec+0.5+DblCalcDev)/10^iDec;
	end
	
end

function CeilDbl(dNum, iDec)

	if iDec == nil then
		iDec = 0;
	end
	if iDec == 0 then
		return math.floor(dNum+1-DblCalcDev);
	else
		return math.floor(dNum*10^iDec+1-DblCalcDev)/10^iDec;
	end
	
end

function FloorDbl(dNum, iDec)

	if iDec == nil then
		iDec = 0;
	end
	if iDec == 0 then
		return math.floor(dNum+DblCalcDev);
	else
		return math.floor(dNum*10^iDec+DblCalcDev)/10^iDec;
	end

end

function IsDbl(dNum1, dNum2)

	return (math.abs(dNum1-dNum2) <= DblCalcDev);

end

function SmallerDbl(dNum1, dNum2)

	return ((dNum2-dNum1) > DblCalcDev);

end

function GreaterDbl(dNum1, dNum2)

	return ((dNum1-dNum2) > DblCalcDev);

end

function MinDbl(dNum1, dNum2)

	if dNum1 <= dNum2 then
		return dNum1;
	else
		return dNum2;
	end

end

function MaxDbl(dNum1, dNum2)

	if dNum1 >= dNum2 then
		return dNum1;
	else
		return dNum2;
	end

end
