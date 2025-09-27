// AUTO GENERATED, DO NOT MODIFY!
// file: test/res/item.xlsx

import {
    BagType,
    ItemType,
    QualityType,
} from "../define/index";

export interface GeneratedItemFollowRow {
    /**
     * ### (location: A1) (checker: x)
     */
    readonly id: number;
    /**
     * 物品类型 (location: B1) (checker: x)
     */
    readonly name?: string;
    /**
     *  (location: C1) (checker: @follow(name))
     */
    readonly value?: string;
}

export interface GeneratedItemItemRow {
    /**
     * ### (location: A1) (checker: x)
     */
    readonly id: number;
    /**
     * 注释 (location: B1) (checker: x)
     */
    readonly comment: string;
    /**
     * 物品名称 (location: C1) (checker: x)
     */
    readonly name: string;
    /**
     * 物品说明 (location: D1) (checker: x)
     */
    readonly desc: string;
    /**
     * 物品类型 config.ITEM_TYPE (location: E1) (checker: x)
     */
    readonly item_type: ItemType;
    /**
     * 背包类型 config.BAG_TYPE (location: F1) (checker: x)
     */
    readonly bag_id: BagType;
    /**
     * 可否堆叠 (location: G1) (checker: x)
     */
    readonly stack?: number;
    /**
     * 品质(颜色) (location: H1) (checker: x)
     */
    readonly quality: QualityType;
    /**
     * 参数 (location: I1) (checker: x)
     */
    readonly args?: unknown;
    /**
     * 背包是否隐藏 (location: J1) (checker: x)
     */
    readonly hide?: boolean;
}

export interface GeneratedItemMapRow {
    /**
     * 注释 (location: B2) (checker: x)
     */
    readonly comment: string;
    /**
     *  (location: C2) (checker: x)
     */
    readonly kind: number;
    /**
     *  (location: D2) (checker: x)
     */
    readonly level: number;
    /**
     * 物品名称 (location: E2) (checker: x)
     */
    readonly name: string;
}

export interface GeneratedItemMapArrRow {
    /**
     * ### (location: A2) (checker: x)
     */
    readonly id: number;
    /**
     * 注释 (location: B2) (checker: x)
     */
    readonly comment: string;
    /**
     *  (location: C2) (checker: x)
     */
    readonly kind: number;
    /**
     *  (location: D2) (checker: x)
     */
    readonly level: number;
    /**
     * 物品名称 (location: E2) (checker: x)
     */
    readonly name: string;
}

export interface GeneratedItemMapFieldRow {
    /**
     * ### (location: A2) (checker: x)
     */
    readonly id: number;
    /**
     * 注释 (location: B2) (checker: x)
     */
    readonly comment: string;
    /**
     *  (location: C2) (checker: x)
     */
    readonly kind: number;
    /**
     *  (location: D2) (checker: x)
     */
    readonly level: number;
    /**
     * 物品名称 (location: E2) (checker: x)
     */
    readonly name: string;
}

export interface GeneratedItemMapObjRow {
    /**
     * ### (location: A2) (checker: x)
     */
    readonly id: number;
    /**
     * 注释 (location: B2) (checker: x)
     */
    readonly comment: string;
    /**
     *  (location: C2) (checker: x)
     */
    readonly kind: number;
    /**
     *  (location: D2) (checker: x)
     */
    readonly level: number;
    /**
     * 物品名称 (location: E2) (checker: x)
     */
    readonly name: string;
}

