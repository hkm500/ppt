"""Compute every number used in the revised deck so nothing is typed from memory."""
import hashlib, math, json

def sha(s): return hashlib.sha256(s.encode()).hexdigest()
def bits_diff(a,b):
    return bin(int(a,16)^int(b,16)).count('1')

out={}
# 1) real SHA-256 examples (replace the fake 128-bit values of the old slides)
ex=["bitcoin","Bitcoin","bitcoin ","morning","hello-1","hello-2","Jack pays Jill 1 BTC","Jack pays Jill 2 BTC"]
out['sha']={e:sha(e) for e in ex}
out['avalanche']={ 'bitcoin vs Bitcoin':bits_diff(sha('bitcoin'),sha('Bitcoin')),
                   'hello-1 vs hello-2':bits_diff(sha('hello-1'),sha('hello-2')),
                   'Jack 1 vs 2':bits_diff(sha('Jack pays Jill 1 BTC'),sha('Jack pays Jill 2 BTC'))}
# double sha of genesis header
hdr=bytes.fromhex("01000000"+"00"*32+"3ba3edfd7a7b12b27ac72c3e67768f617fc81bc3888a51323a9fb8aa4b1e5e4a"+"29ab5f49"+"ffff001d"+"1dac2b7c")
d=hashlib.sha256(hashlib.sha256(hdr).digest()).digest()[::-1].hex()
out['genesis_hash_from_header']=d

# 2) class activity: find nonce so that sha256(text+nonce) starts with N zeros
def mine(prefix,zeros):
    n=0
    while True:
        h=sha(prefix+str(n))
        if h.startswith('0'*zeros): return n,h
        n+=1
for z in (3,4):
    out[f'mine_{z}']=mine("Bitcoin-4-JXUFE-",z)

# 3) white paper section 11: attacker success probability
def P(q,z):
    p=1-q
    lam=z*q/p
    s=1.0
    for k in range(0,z+1):
        poi=math.exp(-lam)
        for i in range(1,k+1): poi*=lam/i
        s-=poi*(1-(q/p)**(z-k))
    return s
out['attack']={str(q):[round(P(q,z),7) for z in range(0,11)] for q in (0.1,0.3)}
# z needed for P<0.1% (white paper)
def zmin(q,thr=0.001):
    z=0
    while P(q,z)>=thr: z+=1
    return z
out['z_for_0.1pct']={str(q):zmin(q) for q in (0.1,0.15,0.2,0.25,0.3,0.35,0.4,0.45)}

# 4) block interval distribution (exponential, mean 10 min)
cdf=lambda t:1-math.exp(-t/10)
bins=[(0,10),(10,20),(20,30),(30,40),(40,1e9)]
out['interval_pct']=[round(100*((cdf(b)-cdf(a)) if b<1e8 else (1-cdf(a))),1) for a,b in bins]
out['gt_10min']=round(100*math.exp(-1),1)
out['gt_60min']=round(100*math.exp(-6),2)

# 5) supply
supply=[];cum=0;r=50;h=0
for i in range(8):
    cum+=210000*r; supply.append((i,r,cum,round(cum/21e6*100,3))); r/=2
out['supply']=supply
out['cap_check']=210000*50*2
# exact cap with satoshi rounding
tot=0;r=50*10**8
while r>0:
    tot+=210000*r; r//=2
out['exact_cap_btc']=tot/1e8
out['annual_issuance_pct_2026']=round(3.125*144*365/20.1e6*100,2)
out['gold_flow_pct']=round(3672/219891*100,2)
out['gold_s2f']=round(219891/3672,1)
out['btc_s2f_post2024']=round(19.7e6/(3.125*144*365),1)
out['btc_s2f_post2020']=round(18.4e6/(6.25*144*365),1)

# 6) scale
out['2pow256']="%.3e"%(2**256)
hps=1e21
out['years_to_bruteforce_256_at_1ZHs']="%.2e"%((2**256)/hps/3.156e7)
out['years_at_exaflop']="%.2e"%((2**256)/1e18/3.156e7)
diff=1.3896e14
out['exp_hashes_per_block']="%.3e"%(diff*2**32)
out['log2_hashes_per_block']=round(math.log2(diff*2**32),1)
out['sec_per_block_at_1ZHs']=round(diff*2**32/1e21)
# merkle proof size
for n in (8,2000,4000):
    out[f'merkle_proof_{n}']=math.ceil(math.log2(n))
print(json.dumps(out,ensure_ascii=False,indent=1))
json.dump(out,open('data.json','w'),ensure_ascii=False,indent=1)
