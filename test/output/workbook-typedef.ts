// AUTO GENERATED, DO NOT MODIFY!

import {
    TCell as _TCell,
} from "../../";
import {
    BagType,
    ItemType,
    Items,
    QualityType,
    TaskType,
} from "./client/define/index";

type TCell = Omit<_TCell, "v">;

// file: test/res/item.xlsx
// processors:
//  - @define
//  - @stringify
//  - @typedef
//  - @auto-register
//  - @post_stringify
//  - @workbook-indexer
//  - @validate-json
//  - @workbook-typedef
export interface ItemDefineRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v?: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    key1: { v: string } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    key2: { v?: string } & TCell;
    /**
     * 注释 (location: E2) (checker: x) (writer: client|server)
     */
    value_comment: { v?: string } & TCell;
    /**
     *  (location: F2) (checker: x) (writer: client|server)
     */
    value: { v: unknown } & TCell;
    /**
     *  (location: G2) (checker: x) (writer: client|server)
     */
    value_type: { v: string } & TCell;
    /**
     *  (location: H2) (checker: x) (writer: client|server)
     */
    enum: { v?: string } & TCell;
}

// file: test/res/item.xlsx
export interface ItemFollowRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 物品类型 (location: B1) (checker: x) (writer: client|server)
     */
    name: { v?: string } & TCell;
    /**
     *  (location: C1) (checker: @follow(name)) (writer: client|server)
     */
    value: { v?: string } & TCell;
}

// file: test/res/item.xlsx
export interface ItemItemRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B1) (checker: x) (writer: client|server)
     */
    comment: { v: string } & TCell;
    /**
     * 物品名称 (location: C1) (checker: x) (writer: client|server)
     */
    name: { v: string } & TCell;
    /**
     * 物品说明 (location: D1) (checker: x) (writer: client)
     */
    desc: { v: string } & TCell;
    /**
     * 物品类型 config.ITEM_TYPE (location: E1) (checker: x) (writer: client|server)
     */
    item_type: { v: ItemType } & TCell;
    /**
     * 背包类型 config.BAG_TYPE (location: F1) (checker: x) (writer: client|server)
     */
    bag_id: { v: BagType } & TCell;
    /**
     * 可否堆叠 (location: G1) (checker: x) (writer: client|server)
     */
    stack: { v?: number } & TCell;
    /**
     * 品质(颜色) (location: H1) (checker: x) (writer: client|server)
     */
    quality: { v: QualityType } & TCell;
    /**
     * 参数 (location: I1) (checker: x) (writer: client|server)
     */
    args: { v?: unknown } & TCell;
    /**
     * 背包是否隐藏 (location: J1) (checker: x) (writer: client|server)
     */
    hide: { v?: boolean } & TCell;
}

// file: test/res/item.xlsx
// processors:
//  - @map(*, kind, level)
export interface ItemMapRow {
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    kind: { v: number } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    level: { v: number } & TCell;
    /**
     * 物品名称 (location: E2) (checker: x) (writer: client|server)
     */
    name: { v: string } & TCell;
}

// file: test/res/item.xlsx
// processors:
//  - @map([name,level], kind, level)
export interface ItemMapArrRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    kind: { v: number } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    level: { v: number } & TCell;
    /**
     * 物品名称 (location: E2) (checker: x) (writer: client|server)
     */
    name: { v: string } & TCell;
}

// file: test/res/item.xlsx
// processors:
//  - @map(.comment, kind, level)
export interface ItemMapFieldRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    kind: { v: number } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    level: { v: number } & TCell;
    /**
     * 物品名称 (location: E2) (checker: x) (writer: client|server)
     */
    name: { v: string } & TCell;
}

// file: test/res/item.xlsx
// processors:
//  - @map({name,level,kind}, kind, level)
export interface ItemMapObjRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    kind: { v: number } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    level: { v: number } & TCell;
    /**
     * 物品名称 (location: E2) (checker: x) (writer: client|server)
     */
    name: { v: string } & TCell;
}

// file: test/res/task.xlsx
export interface TaskBranchRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 任务组 (location: C1) (checker: x) (writer: client|server)
     */
    group: { v: number } & TCell;
    /**
     * 显示排序 (location: D1) (checker: x) (writer: client|server)
     */
    sort: { v: number } & TCell;
    /**
     * 任务类型 (location: E1) (checker: x) (writer: client|server)
     */
    type: { v: TaskType } & TCell;
    /**
     * 名字 (location: F1) (checker: x) (writer: client)
     */
    name: { v: string } & TCell;
    /**
     * 描述 (location: G1) (checker: x) (writer: client|server)
     */
    desc: { v: string } & TCell;
    /**
     * 后置任务 (location: H1) (checker: #branch.id) (writer: client|server)
     */
    next_task: { v?: number } & TCell;
    /**
     * 条件 (location: I1) (checker: x) (writer: client|server)
     */
    condition: { v: string } & TCell;
    /**
     * 累计 (location: J1) (checker: x) (writer: client|server)
     */
    total: { v?: boolean } & TCell;
    /**
     * 参数 (location: K1) (checker: x) (writer: client|server)
     */
    args: { v: unknown } & TCell;
    /**
     * 参数类型 (location: L1) (checker: x) (writer: client|server)
     */
    args_type: { v: string } & TCell;
    /**
     * 奖励 (location: M1) (checker: $[*].id==item#item.id) (writer: client|server)
     */
    reward: { v: Items } & TCell;
    /**
     * 任务icon (location: N1) (checker: x) (writer: client)
     */
    icon: { v: string } & TCell;
    /**
     *  (location: O1) (checker: x) (writer: client)
     */
    custom: { v?: string } & TCell;
    /**
     * 完成后自动提交 (location: P1) (checker: x) (writer: client)
     */
    auto_submit: { v?: number } & TCell;
}

// file: test/res/task.xlsx
// processors:
//  - @config
export interface TaskConfRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 键 (location: B2) (checker: x) (writer: client|server)
     */
    key: { v: string } & TCell;
    /**
     * 值 (location: C2) (checker: x) (writer: client|server)
     */
    value: { v: string } & TCell;
    /**
     * 值类型 (location: D2) (checker: x) (writer: client|server)
     */
    value_type: { v: string } & TCell;
    /**
     * 描述 (location: E2) (checker: x) (writer: client|server)
     */
    value_comment: { v: string } & TCell;
}

// file: test/res/task.xlsx
// processors:
//  - @define
//  - @stringify(task)
//  - @typedef
//  - @auto-register
//  - @post_stringify
//  - @workbook-indexer
//  - @validate-json
//  - @workbook-typedef
export interface TaskDefineRow {
    /**
     * ### (location: A2) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 注释 (location: B2) (checker: x) (writer: client|server)
     */
    comment: { v?: string } & TCell;
    /**
     *  (location: C2) (checker: x) (writer: client|server)
     */
    key1: { v: string } & TCell;
    /**
     *  (location: D2) (checker: x) (writer: client|server)
     */
    key2: { v?: string } & TCell;
    /**
     * 注释 (location: E2) (checker: x) (writer: client|server)
     */
    value_comment: { v?: string } & TCell;
    /**
     *  (location: F2) (checker: $&key2=MAIN==#main.type&condition=mainline_event) (writer: client|server)
     */
    value: { v: unknown } & TCell;
    /**
     *  (location: G2) (checker: x) (writer: client|server)
     */
    value_type: { v: string } & TCell;
    /**
     *  (location: H2) (checker: x) (writer: client|server)
     */
    enum: { v?: string } & TCell;
    /**
     *  (location: I2) (checker: x) (writer: client|server)
     */
    enum_option: { v?: boolean } & TCell;
}

// file: test/res/task.xlsx
export interface TaskEventsRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 显示排序 (location: C1) (checker: x) (writer: client|server)
     */
    sort: { v: number } & TCell;
    /**
     * 任务类型 (location: D1) (checker: x) (writer: client|server)
     */
    type: { v: number } & TCell;
    /**
     * 描述 (location: E1) (checker: x) (writer: client|server)
     */
    desc: { v: string } & TCell;
    /**
     * 条件 (location: F1) (checker: x) (writer: server)
     */
    condition: { v: string } & TCell;
    /**
     * 参数 (location: G1) (checker: x) (writer: server)
     */
    args: { v: unknown } & TCell;
    /**
     * 奖励 (location: H1) (checker: $[*].id==item#item.id) (writer: client|server)
     */
    reward: { v: Items } & TCell;
    /**
     * vip奖励 (location: I1) (checker: $[*].id==item#*.id) (writer: client|server)
     */
    vip_reward: { v: Items } & TCell;
    /**
     * 任务icon (location: J1) (checker: x) (writer: client)
     */
    icon: { v: string } & TCell;
}

// file: test/res/task.xlsx
export interface TaskExchangeRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 显示排序 (location: C1) (checker: x) (writer: client|server)
     */
    sort: { v: number } & TCell;
    /**
     * 任务类型 (location: D1) (checker: x) (writer: client|server)
     */
    type: { v: number } & TCell;
    /**
     * 描述 (location: E1) (checker: x) (writer: client|server)
     */
    desc: { v: string } & TCell;
    /**
     * 后置任务 (location: F1) (checker: task#*.id) (writer: client|server)
     */
    next_task: { v?: number } & TCell;
    /**
     * 条件 (location: G1) (checker: x) (writer: server)
     */
    condition: { v: string } & TCell;
    /**
     * 累计 (location: H1) (checker: x) (writer: client|server)
     */
    total: { v?: boolean } & TCell;
    /**
     * 参数 (location: I1) (checker: x) (writer: client|server)
     */
    args: { v: unknown } & TCell;
    /**
     * 奖励 (location: J1) (checker: @ItemArrayChecker) (writer: client|server)
     */
    reward: { v: Items } & TCell;
    /**
     * 任务icon (location: K1) (checker: x) (writer: client)
     */
    icon: { v: string } & TCell;
    /**
     * 任务npc和对话 (location: L1) (checker: x) (writer: client)
     */
    custom: { v?: unknown } & TCell;
}

// file: test/res/task.xlsx
export interface TaskMainRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 显示排序 (location: C1) (checker: $ > 0 && $ < 20) (writer: client|server)
     */
    sort: { v: number } & TCell;
    /**
     * 任务类型 (location: D1) (checker: #define.value&key1=TASK_TYPE) (writer: client|server)
     */
    type: { v: TaskType } & TCell;
    /**
     * 描述 (location: E1) (checker: x) (writer: client|server)
     */
    desc: { v: string } & TCell;
    /**
     * 后置任务 (location: F1) (checker: task#*.id&type=MAIN) (writer: client|server)
     */
    next_task: { v?: number } & TCell;
    /**
     * 支线任务 (location: G1) (checker: $[*]==#branch.id) (writer: client|server)
     */
    branch_tasks: { v?: number[] } & TCell;
    /**
     * 条件 (location: H1) (checker: x) (writer: client|server)
     */
    condition: { v: string } & TCell;
    /**
     * 累计 (location: I1) (checker: x) (writer: client|server)
     */
    total: { v?: boolean } & TCell;
    /**
     * 参数 (location: J1) (checker: @TaskArgsChecker) (writer: client|server)
     */
    args: { v: unknown } & TCell;
    /**
     * 奖励 (location: K1) (checker: @ItemArrayChecker) (writer: client|server)
     */
    reward: { v: Items } & TCell;
    /**
     * 任务icon (location: L1) (checker: x) (writer: client)
     */
    icon: { v: string } & TCell;
    /**
     *  (location: M1) (checker: x) (writer: client)
     */
    custom: { v?: unknown } & TCell;
    /**
     * 完成后自动提交 (location: N1) (checker: [1]) (writer: client)
     */
    auto_submit: { v?: number } & TCell;
}

// file: test/res/task.xlsx
export interface TaskWeeklyRow {
    /**
     * ### (location: A1) (checker: x) (writer: client|server)
     */
    id: { v: number } & TCell;
    /**
     * 显示排序 (location: C1) (checker: x) (writer: client|server)
     */
    sort: { v: number } & TCell;
    /**
     * 任务类型 (location: D1) (checker: x) (writer: client|server)
     */
    type: { v: number } & TCell;
    /**
     * 描述 (location: E1) (checker: x) (writer: client|server)
     */
    desc: { v: string } & TCell;
    /**
     * 条件 (location: F1) (checker: x) (writer: server)
     */
    condition: { v: string } & TCell;
    /**
     * 参数 (location: G1) (checker: x) (writer: server)
     */
    args: { v: unknown } & TCell;
    /**
     * 奖励 (location: H1) (checker: @ItemArrayChecker) (writer: client|server)
     */
    reward: { v: Items } & TCell;
    /**
     * 任务icon (location: I1) (checker: x) (writer: client)
     */
    icon: { v: string } & TCell;
}


