# XLSX Exporter

一个强大的 Excel 文件转换工具，支持将 Excel 文件转换为多种编程语言的配置文件和数据文件。

## 目录

- [功能特性](#功能特性)
- [安装](#安装)
- [快速开始](#快速开始)
    - [Excel 文件格式说明](#excel-文件格式说明)
    - [基本用法](#基本用法)
    - [完整示例](#完整示例)
- [数据类型](#数据类型)
- [数据验证系统](#数据验证系统)
    - [内置检查器](#内置检查器)
    - [检查器语法](#检查器语法)
    - [高级索引检查器](#高级索引检查器)
- [检查器详细语法](#检查器详细语法)
    - [表格结构说明](#表格结构说明)
    - [核心操作符说明](#核心操作符说明)
    - [语法形式](#语法形式)
    - [行键表达式语法](#行键表达式语法)
    - [过滤器语法](#过滤器语法)
    - [使用示例](#使用示例)
    - [常见应用场景](#常见应用场景)
    - [语法规则总结](#语法规则总结)
- [输出格式](#输出格式)
- [高级功能](#高级功能)
- [开发和调试](#开发和调试)
- [许可证](#许可证)
- [支持](#支持)

## 功能特性

- 📊 **Excel 解析** - 支持解析.xlsx 格式的 Excel 文件
- 🔄 **多语言输出** - 支持输出 TypeScript、Lua、JSON 等格式
- 🏷️ **类型定义生成** - 自动生成 TypeScript 和 Lua 的类型定义
- ✅ **数据验证** - 内置数据检查器确保数据质量
- 🔧 **可扩展架构** - 支持自定义转换器、检查器和处理器
- 📝 **注释支持** - 保留 Excel 中的注释信息
- 🎯 **配置驱动** - 通过 Excel 表格配置数据结构

## 安装

```bash
npm i
```

## 快速开始

### Excel 文件格式说明

Excel 文件需要按照特定格式组织：

- 第 1 行是配置表格处理器（可选）
- 第 2 行是字段名：id, comment ...
- 第 3 行是数据类型：int, float, string, string? ...
- 第 4 行是导出：x 代表不导出，client 只导客户端，server 只导服务端，不填默认导出双端
- 第 5 行是检查：x 代表不检查，可用的检查器见下面
- 第 6 行是注释

<table style="text-align: center;">
  <tr>
    <td colspan="8" style="text-align: left; background-color: #f0f0f0; font-weight: bold;">@config;@stringify(merge);@typedef</td>
  </tr>
  <tr>
    <th>id</th>
    <th>name</th>
    <th>desc</th>
    <th>positon</th>
    <th>item</th>
    <th>task</th>
    <th>reward</th>
    <th>cond</th>
  </tr>
  <tr>
    <td>int</td>
    <td>string</td>
    <td>string?</td>
    <td>int</td>
    <td>int</td>
    <td>int</td>
    <td>int</td>
    <td>int</td>
  </tr>
  <tr>
    <td>>></td>
    <td>client</td>
    <td>server</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td>!!!</td>
    <td>@TreeChker</td>
    <td>x</td>
    <td>[0,1,2]</td>
    <td>item#item.id</td>
    <td>task#*.id</td>
    <td>id==reward#*.id<br />@CheckReward</td>
    <td>#level.id</td>
  </tr>
  <tr>
    <td>###</td>
    <td>名字</td>
    <td>描述</td>
    <td>位置</td>
    <td>物品</td>
    <td>任务</td>
    <td>奖励</td>
    <td>条件</td>
  </tr>
  <tr>
    <td>1</td>
    <td>关哥</td>
    <td>大长枪</td>
    <td>1</td>
    <td>{101, 103}</td>
    <td>666</td>
    <td>{id = 3}</td>
    <td>1</td>
  </tr>
  <tr>
    <td>2</td>
    <td>张飞</td>
    <td>大长枪</td>
    <td>1</td>
    <td>101</td>
    <td>{666}</td>
    <td>{id=4}</td>
    <td>{1， 2， 3}</td>
  </tr>
</table>

### 基本用法

```typescript
import * as xlsx from "xlsx-exporter";

// 注册输出适配器
xlsx.registerWriter("client", (path, data, processor) => {
    // 处理客户端输出
});

xlsx.registerWriter("server", (path, data, processor) => {
    // 处理服务端输出
});

// 解析Excel文件
xlsx.parse(["path/to/your/file.xlsx"]);
```

### 完整示例

```typescript
import * as xlsx from "xlsx-exporter";
import { toPascalCase } from "./src/util";

const OUTPUT_DIR = "output";

// 注册客户端输出适配器
xlsx.registerWriter("client", (path, data, processor) => {
    if (processor === "config") {
        // 生成TypeScript配置文件
        const name = toPascalCase(`${xlsx.filename(path)}_${data["!name"]}`);
        const marshal = `export const ${name} = `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/client/config/${name}.ts`,
            xlsx.stringifyTs(data, { indent: 4, marshal })
        );
    } else if (processor === "stringify") {
        // 生成JSON数据文件
        const name = xlsx.filename(path);
        xlsx.writeFile(
            `${OUTPUT_DIR}/client/data/${name}.json`,
            xlsx.stringifyJson(data, { indent: 2 })
        );
    } else if (processor === "gen-type") {
        // 生成TypeScript类型定义
        const name = xlsx.filename(path);
        const types = xlsx.genTsType(path, "client");
        const content = `// AUTO GENERATED, DO NOT MODIFY!\n\n${types}`;
        xlsx.writeFile(`${OUTPUT_DIR}/client/types/${name}.ts`, content);
    }
});

// 注册服务端输出适配器
xlsx.registerWriter("server", (path, data, processor) => {
    if (processor === "config") {
        // 生成Lua配置文件
        const name = `${xlsx.filename(path)}_${data["!name"]}`;
        const marshal = `return `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/server/config/${name}.lua`,
            xlsx.stringifyLua(data, { indent: 4, marshal })
        );
    } else if (processor === "stringify") {
        // 生成Lua数据文件
        const name = xlsx.filename(path);
        const marshal = `return `;
        xlsx.writeFile(
            `${OUTPUT_DIR}/server/data/${name}.lua`,
            xlsx.stringifyLua(data, { indent: 2, marshal })
        );
    } else if (processor === "gen-type") {
        // 生成Lua类型定义
        const name = xlsx.filename(path);
        const types = xlsx.genLuaType(path, "server");
        const content = `-- AUTO GENERATED, DO NOT MODIFY!\n\n${types}`;
        xlsx.writeFile(`${OUTPUT_DIR}/server/types/${name}.lua`, content);
    }
});

// 解析Excel文件
xlsx.parse(["data/item.xlsx", "data/task.xlsx"]);
```

## 数据类型

- `int` - 整数
- `float` - 浮点数
- `string` - 字符串
- `bool` - 布尔值
- `json` - JSON 对象
- `table` - lua 表格数据
- `auto` - 自动类型推断

## 数据验证系统

### 内置检查器

- `size` - 数据大小检查
- `range` - 数值范围检查
- `index` - 索引验证
- `follow` - 依赖关系检查
- `expr` - 表达式验证

### 检查器语法

- #skill.id
  当前列的值必须在当前表的 `skill` 分页中的 `id` 字段中找到
- battle/battle_skill#skill.id
  当前列的值必须在 `battle_skill` 表的 `skill` 分页中的 `id` 字段中找到
- battle/battle_skill#\*.id
  当前列的值必须在 `battle_skill` 表的任意分页中的 `id` 字段中找到
- battle/battle_skill#skill.id=id
  当前列的子字段 `id` 值必须在 `battle_skill` 表的 `skill` 分页中的 `id` 字段中找到
- \$ >= 1 && \$ <= 9
  表达式检查，示例在 recharge.xlsx 的 discount 字段
- [1,2,3,4]
  检查值在不在这个数组内
- @size(10) 表示字段要是 table 属性，而且长度只能为 10
  检查 table 字段的长度
- @follow(field)
  表示当前单元格有没有值要与目标列一致。

### 高级索引检查器

**核心机制**：

- `#` 是"取表"操作符，用于指定目标表格位置
- 根据是否有行表达式、行过滤器或列过滤器来选择语法形式

## 检查器详细语法

### 表格结构说明

基于项目中Excel文件的标准结构：

```
第1行: @define;@stringify(表名)           // 处理器定义
第2行: id | comment | key1 | key2 | ...   // 字段名
第3行: int | string? | string | ...        // 字段类型
第4行: >> |   |   |   | ...               // 可选的状态标记
第5行: !!! | x | x | x | ...             // 必填字段标记
第6行: ### | 注释 |   |   | ...           // 字段注释
第7行开始: 实际数据
```

### 语法形式

**核心操作符说明**：

- **`#`** 是"取表"操作符，用于指定目标表格位置
- **`==`** 是分隔符，在特定情况下使用

#### 1. 简单形式（直接检查当前单元格值）

```
[文件名]#[工作表名].[列名]
```

#### 2. 带列过滤器形式（左边有筛选时，左边必须有$）

```
$[表达式]==[文件名]#[工作表名].[列名]&[列过滤器]
```

#### 3. 完整形式（有行表达式或行过滤器）

```
$[行键表达式][&行过滤器]==[文件名]#[工作表名].[列名][&列过滤器]
```

**文件名规则**：

- **当前文件内查找**：可以省略文件名，如 `#hero.id`
- **跨文件引用**：必须指定文件名，如 `hero#hero.id`

**关键规则**：

- **左边有过滤器时**：左边必须要有 `$` 表达式，使用 `==` 分隔
- **有行表达式或行过滤器时**：使用 `==` 分隔
- **简单检查当前单元格值时**：直接使用 `#` 取表操作符

### 行键表达式语法

#### 重要说明

在行键表达式中，`$` 代表**当前单元格的值**，而不是当前行的值。这意味着：

- 如果当前单元格包含简单值（如数字、字符串），则 `$` 就是该值
- 如果当前单元格包含JSON对象，则可以用 `$.property` 访问对象属性
- 如果当前单元格包含数组，则可以用 `$[index]` 访问数组元素

#### 基本路径

- `.property` - 访问对象属性
- `[index]` - 访问数组元素（从0开始）
- `[*]` - 遍历数组所有元素
- `[.]` - 获取对象所有键名

#### 可选访问

在路径后加 `?` 表示可选访问，如果路径不存在则跳过而不报错：

- `.property?` - 可选属性访问
- `[index]?` - 可选数组元素访问

#### 复杂路径示例

- `$.id` - 获取当前单元格值（如果是对象）的id属性
- `$.rewards[*].item_id` - 获取当前单元格值中rewards数组所有元素的item_id
- `$.config.targets[0]` - 获取当前单元格值中config对象的targets数组第一个元素

### 过滤器语法

过滤器使用 `&` 连接多个条件，格式为 `字段名=值`：

- `type=MAIN` - 当前行的type字段等于MAIN
- `quality=1&enabled=true` - 当前行的quality字段为1且enabled字段为true

**注意**：过滤器中的 `=` 是单等号，用于字段匹配；而 `==` 是双等号，用于分隔整个检查表达式的左右两部分。

### 使用示例

#### 基于项目实际案例的示例

以下示例均来自项目中 `design-types.ts`的真实checker规则：

#### 示例1：简单索引检查

```yaml
# 检查功能开启ID是否存在
# 来源：activity.xlsx -> activity工作表
func_id: open_func#func.id

# 检查英雄ID是否存在
# 来源：battle/battle_robot.xlsx -> hero工作表
hero_id: hero#hero.id

# 检查怪物ID是否存在
# 来源：activity/battle_pass.xlsx -> monster工作表
monster_id: monster#troop.id

# 检查价格是否在价格表中存在
# 来源：activity/accumulate_recharge.xlsx -> reward工作表
cost: price#price.cny
```

#### 示例2：带列过滤器的检查

```yaml
# 检查装备ID是否在对应部位的装备中存在
# 来源：battle/battle_test.xlsx -> t1工作表
eq_part_1: $==equipment#equipment.id&part=1 # 头盔
eq_part_2: $==equipment#equipment.id&part=2 # 战甲
eq_part_6: $==equipment#equipment.id&part=6 # 武器

# 检查联盟道具购买价格中的道具ID
# 来源：alliance.xlsx -> item工作表
buy_price: $[*].id==#item.id
```

#### 示例3：带行过滤器的检查

```yaml
# 只有当key1为COLLECTION_ITEM_ID时才检查物品ID
# 来源：activity/wusheng_road.xlsx -> define工作表
value: $&key1=COLLECTION_ITEM_ID==item#item.id

# 根据不同条件检查不同表（多条件可选验证）
# 来源：activity/upstar_limit.xlsx -> task工作表
args: $.star?==hero#hero_star.star;$.stage?==hero#hero_stage.stage_parameter
```

#### 示例4：数组元素检查

```yaml
# 检查任务数组中每个ID是否都存在
# 来源：activity/battle_pass.xlsx -> typeInfo工作表
daily_tasks: $[*]==activity/battle_pass#task.task_id
weekly_tasks: $[*]==activity/battle_pass#task.task_id

# 检查技能动作ID数组
# 来源：battle/battle_skill.xlsx -> skill工作表
carry_actions: $[*]==battle/battle_skill#action.id

# 检查技能标签数组
# 来源：battle/battle_skill.xlsx -> buff工作表
granted_tags: $[*]==#define.key2&key1=SKILL_TAG
```

#### 示例5：对象键检查

```yaml
# 检查前置科技条件（对象的键）
# 来源：alliance.xlsx -> technology工作表
pre_tech_cond: $[.]==#technology.tech_id
```

#### 示例6：条件性检查

```yaml
# 根据不同属性检查不同表（可选属性验证）
# 来源：activity/upstar_limit.xlsx -> task工作表
args: $.star?==hero#hero_star.star;$.stage?==hero#hero_stage.stage_parameter

# 复杂的属性检查（多层可选验证）
# 来源：alliance.xlsx -> technology工作表
base: $.higner_attrs?[*][0]==attr#higher_attr.id;$.attrs?[*][0]==attr#attr.id
```

#### 示例7：跨目录文件引用

```yaml
# 检查传送点奖励
# 来源：activity/novice_limit_time.xlsx -> drop工作表
transferId: battle/battle_pve_map#transfer.id

# 检查NPC状态
# 来源：battle/battle_npc_state.xlsx -> npcState工作表
npc_id: battle/battle_npc#npc.id

# 检查获取途径ID
# 来源：activity/battle_pass.xlsx -> task工作表
getwayid: item#itemGetWay.id
```

#### 示例8：复杂嵌套检查

```yaml
# 检查属性数组，每个元素的第一个值必须是属性ID
# 来源：battle/battle_skill_lv.xlsx -> attr工作表
attr: $[*][0]==attr#attr.id

# 检查任务ID（支持通配符）
# 来源：battle/battle_interaction_resource.xlsx -> resource工作表
born_task_id: task#*.id

# 检查资产ID
# 来源：alliance.xlsx -> building工作表
asset_id: asset#assets.id
```

### 常见应用场景

#### 1. 外键关系验证

最常见的用法，验证ID字段的外键关系：

```yaml
# 活动功能开启检查
# 来源：activity.xlsx -> activity工作表
func_id: open_func#func.id

# 英雄ID验证
# 来源：battle/battle_robot.xlsx -> hero工作表
hero_id: hero#hero.id

# 怪物ID验证（跨文件）
# 来源：activity/battle_pass.xlsx -> monster工作表
monster_id: monster#troop.id

# 资产ID验证
# 来源：alliance.xlsx -> building工作表
asset_id: asset#assets.id
```

#### 2. 带条件的验证

根据其他字段值进行条件性检查：

```yaml
# 装备部位验证：根据装备部位检查对应的装备
# 来源：battle/battle_test.xlsx -> t1工作表
eq_part_1: $==equipment#equipment.id&part=1 # 头盔
eq_part_6: $==equipment#equipment.id&part=6 # 武器

# 价格验证：检查价格是否在价格表中存在
# 来源：activity/daily_recharge.xlsx -> reward工作表
recharge_limit: price#price.cny
# 来源：activity/gift_push.xlsx -> gifts工作表
cost: price#price.cny
```

#### 3. 数组和集合验证

验证数组中每个元素或对象的键：

```yaml
# 任务列表验证
# 来源：activity/battle_pass.xlsx -> typeInfo工作表
daily_tasks: $[*]==activity/battle_pass#task.task_id
weekly_tasks: $[*]==activity/battle_pass#task.task_id

# 技能动作验证
# 来源：battle/battle_skill.xlsx -> skill工作表
carry_actions: $[*]==battle/battle_skill#action.id

# 对象键验证（前置科技）
# 来源：alliance.xlsx -> technology工作表
pre_tech_cond: $[.]==#technology.tech_id

# 属性数组验证（数组元素的第一个值）
# 来源：battle/battle_skill_lv.xlsx -> attr工作表
attr: $[*][0]==attr#attr.id
```

#### 4. 复杂条件验证

根据行过滤器进行复杂的条件验证：

```yaml
# 根据key1字段值决定是否验证
# 来源：activity/wusheng_road.xlsx -> define工作表
value: $&key1=COLLECTION_ITEM_ID==item#item.id

# 标签验证：根据标签类型进行验证
# 来源：battle/battle_skill.xlsx -> buff工作表
granted_tags: $[*]==#define.key2&key1=SKILL_TAG
ongoing_require_tags: $[*]==#define.key2&key1=SKILL_TAG
```

#### 5. 多条件可选验证

使用 `?`进行可选字段验证：

```yaml
# 根据不同属性检查不同表
# 来源：activity/upstar_limit.xlsx -> task工作表
args: $.star?==hero#hero_star.star;$.stage?==hero#hero_stage.stage_parameter

# 复杂属性验证
# 来源：alliance.xlsx -> technology工作表
base: $.higner_attrs?[*][0]==attr#higher_attr.id;$.attrs?[*][0]==attr#attr.id
percent: $.higner_attrs?[*][0]==attr#higher_attr.id;$.attrs?[*][0]==attr#attr.id
```

#### 6. 跨目录文件验证

验证不同子目录中的表格引用：

```yaml
# 战斗相关验证
# 来源：activity/novice_limit_time.xlsx -> drop工作表
transferId: battle/battle_pve_map#transfer.id
# 来源：battle/battle_npc_state.xlsx -> npcState工作表
npc_id: battle/battle_npc#npc.id

# 活动相关验证
# 来源：activity/battle_pass.xlsx -> task工作表
getwayid: item#itemGetWay.id

# 技能相关验证
# 来源：battle/battle_test.xlsx -> ft1工作表
skill1_id: battle/battle_skill#skill.id
```

#### 7. 通配符表名验证

使用通配符匹配多个工作表：

```yaml
# 支持任意工作表的任务ID
# 来源：battle/battle_interaction_resource.xlsx -> resource工作表
born_task_id: task#*.id

# 支持任意工作表的功能ID
# 来源：activity/fund.xlsx -> fundInfo工作表
func_jump: open_func#*.id
```

### 语法规则总结

基于项目实际使用情况的完整语法总结：

#### 基本规则

- **`#` 是"取表"操作符**：指定目标表格
- **文件名可省略**：当前文件内用 `#表名.列名`，跨文件用 `文件名#表名.列名`
- **支持子目录**：如 `battle/battle_skill#skill.id`
- **支持通配符**：如 `task#*.id`（匹配任意工作表）

#### 路径表达式语法

`$` - 当前单元格值

`$.property` - 对象属性

`$[index]` - 数组元素

`$[*]` - 数组所有元素

`$[.]` - 对象所有键

`$.property?` - 可选属性（不存在时跳过）

`$[*][0]` - 数组元素的第一个值

#### 实际使用模式

```yaml
# 模式1：简单ID验证
hero_id: hero#hero.id

# 模式2：带列过滤器的验证
eq_part_1: $==equipment#equipment.id&part=1

# 模式3：数组元素验证
tasks: $[*]==activity/battle_pass#task.task_id

# 模式4：对象键验证
tech_cond: $[.]==#technology.tech_id

# 模式5：条件验证
value: $&key1=ITEM_ID==item#item.id

# 模式6：可选属性验证
args: $.star?==hero#hero_star.star

# 模式7：跨目录验证
npc_id: battle/battle_npc#npc.id
```

#### !@checker

所有检查器前缀带 “!”，就表明，不管当前单元格有没有值，都要执行检查。

## 输出格式

#### JSON & TypeScript 输出

```typescript
// 配置文件
export const ItemConfig = {
    BAG_TYPE: BagType,
    FIRST_ITEM_UID: 1000000,
    ITEM_TYPE: ItemType,
    QUALITY: QualityType,
}

// 数据文件
{
  "10101": {
    "id": 10101,
    "name": "普通宝箱",
    "quality": 1,
    "stack": 1
  }
}
```

#### Lua 输出

```lua
-- 配置文件
return {
    BAG_TYPE = {
        ITEM = 1,
        CHEST = 2,
    },
    FIRST_ITEM_UID = 1000000,
}

-- 数据文件
return {
    ["10101"] = {
        id = 10101,
        name = "普通宝箱",
        quality = 1,
        stack = 1
    }
}
```

## 高级功能

### 自定义类型转换器

```typescript
// 注册自定义类型
xlsx.registerType("custom", (str: string) => {
    // 自定义转换逻辑
    return customConvert(str);
});
```

### 自定义数据检查器

```typescript
// 注册自定义检查器
xlsx.registerChecker("CustomChecker", (...args: string[]) => {
    return (cell, row, field, errors) => {
        // 自定义验证逻辑
        return true;
    };
});
```

### 自定义处理器

```typescript
// 注册自定义处理器
xlsx.registerProcessor(
    "CustomProcessor",
    (workbook, sheet, ...args) => {
        // 自定义处理逻辑
    },
    100
); // 优先级
```

## 开发和调试

运行测试：

```bash
npm run test
```

## 许可证

MIT License

## 支持

如果您遇到问题或有建议，请在 GitHub 上提交 issue。
