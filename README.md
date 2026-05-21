# PPT Optimizer

## 涓枃璇存槑

杩欐槸涓€涓湰鍦拌繍琛岀殑 PPT 浼樺寲鍛戒护琛屽伐鍏凤紝鐢ㄤ簬鍒嗘瀽鍜屼紭鍖?`.pptx`
鏂囦欢銆傚畠鐩存帴澶勭悊 PowerPoint 鐨?Open XML 鍖呯粨鏋勶紝涓嶄緷璧?PowerPoint
鎴?`python-pptx`銆傚鏋滄湰鏈哄畨瑁呬簡 Pillow锛岃繕鍙互瀵?PPT 鍐呯殑澶у浘鐗囪繘琛屽帇缂┿€?
### 鍔熻兘

- 鐢熸垚 PPT 鍒嗘瀽鎶ュ憡銆?- 缁熶竴鎸囧畾椤甸潰鎴栨暣浠?PPT 鐨勫瓧浣撱€?- 鍒犻櫎婕旇鑰呭娉ㄥ拰璇勮鐩稿叧淇℃伅銆?- 鍘嬬缉鎴栫缉鏀?PPT 鍐呰繃澶х殑鍥剧墖銆?- 杈撳嚭鏂扮殑浼樺寲鍓湰锛屼笉浼氳鐩栧師鏂囦欢銆?- 鏀寔鍙紭鍖栨寚瀹氶〉锛屼緥濡傜 3 椤垫垨绗?2銆?-7 椤点€?
### 浣跨敤鏂规硶

鍏堣繘鍏ラ」鐩洰褰曪細

```powershell
cd C:\Users\22625\Documents\Playground
```

浼樺寲鏁翠唤 PPT锛?
```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --font "Microsoft YaHei"
```

鍙垎鏋愶紝涓嶅鍑烘柊鏂囦欢锛?
```powershell
python -m ppt_optimizer input.pptx --report-only
```

鍙紭鍖栨煇涓€椤碉細

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 3
```

鍙紭鍖栧椤碉細

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 2,5-7
```

鍘嬬缉鍥剧墖锛?
```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --image-quality 72 --max-image-width 1920
```

### 璇存槑

褰撳墠浼樺寲鍣ㄥ亸淇濆畧锛氬畠浼氬敖閲忎繚鐣欏師 PPT 鐨勫竷灞€銆佸姩鐢汇€佸浘琛ㄥ拰姣嶇増缁撴瀯锛屽彧淇敼瀛椾綋銆?澶囨敞銆佽瘎璁恒€佸浘鐗囧拰鐩稿叧 Open XML 鍖呯粨鏋勩€傞€傚悎鍏堝仛鍗曢〉瀹氬悜浼樺寲锛屽啀閫愭鎵╁睍鍒版洿澶嶆潅鐨?瑙嗚閲嶆帓鑳藉姏銆?
## English

A small local CLI for analyzing and optimizing `.pptx` files.

It works directly with the PowerPoint Open XML package, so it does not require
PowerPoint or `python-pptx`. If Pillow is installed, it can also recompress large
JPEG/PNG images.

## Features

- Generate a slide-by-slide report.
- Normalize fonts across text runs.
- Remove speaker notes and comment authors.
- Recompress oversized images when Pillow is available.
- Write a new optimized copy without changing the original file.

## Usage

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --font "Microsoft YaHei"
```

Analyze without writing a new file:

```powershell
python -m ppt_optimizer input.pptx --report-only
```

More aggressive image compression:

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --image-quality 72 --max-image-width 1920
```

Optimize only selected slides:

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 3
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 2,5-7
```

## Notes

The optimizer is intentionally conservative. It keeps layout, animations, charts,
and slide masters intact while changing only scoped XML/package parts.
