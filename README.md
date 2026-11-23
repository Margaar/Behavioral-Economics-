FTC Δ Font Color =
SWITCH (
    TRUE (),
    ISBLANK ( [FTC Δ] ), BLANK (),
    [FTC Δ] > 0, "#1E7B34",     -- зелёный
    [FTC Δ] < 0, "#C00000",     -- красный
    "#666666"                   -- 0 = серый
)

FTC Δ Bg Color =
SWITCH (
    TRUE (),
    ISBLANK ( [FTC Δ] ), BLANK (),
    [FTC Δ] > 0, "#E6F4EA",     -- светло-зелёный фон
    [FTC Δ] < 0, "#FDE7E7",     -- светло-красный фон
    BLANK ()
)
