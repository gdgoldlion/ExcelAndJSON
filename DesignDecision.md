ExcelAndJSON的设计决策
============

*很多人看到ExcelAndJSON的第一反映是，这东西我的公司里面也有，那么我为什么用呢？*

*做为开发来说，每一个工具的存在，都是为了加快游戏开发的速度。那么从无到有，从有到精。有和没有，好用和不好用的差别，每一个都比前一个情况能提升50%的效率。（按IPD理论，极限速度是提升100%的效率，这里取保守数字）。*

*ExcelAndJSON这个工具，前前后后设计思考大约有一年的时间。在之前的开发中，我们使用大量的类似工具，数量有四五个，如果考虑评估阶段的话，是十几个。*

*这些工具或多或少都有这样那样的问题。而每一个问题，都是开发中的一个大坑。下面我们来看ExcelAndJSON是如何对这些问题提供解决方案的。*

**Part1.为什么选择Python开发？**
============

- 如果选择C++，那么是可以使用Qt的。但C++领域，一直没有好用的开源跨平台Excel解析库。要么是闭源的，要么是只支持老格式xls不支持新格式xlsx，还有就是不能够跨平台。而这些，恰恰都是ExcelAndJSON本身必须具备的特性。手游开发，决定了必须跨平台。开源项目决定了依赖库必须也是开源的。Office的不断更新，决定了必须支持新格式。所以，C++被淘汰出局。

- 如果选择JS，因为我的方向是全栈式，目前来说在Node.js领域，npm中我没有找到非常好用的Excel解析库。很多库都是直接把Excel读成一个巨型JSON对象，这种写法是我所不能接受的，太SB了。还有一个原因在于，考虑未来扩展性，Node.js领域一直没有好用的UI库。另外，如果在web开发里面去找，我个人不是很喜欢BS架构的工具。所以，JS被淘汰出局。

- 如果选择Java。Java目前在前端手机游戏开发领域，已经没落。在后端，快速开发方向面临新兴方案的冲击（RoR， Python，Node.js，Go），而且高性能方向又始终干不过C++。对于各个公司自行修改维护能否找到适合的人，是个问题（前端几乎没人做Java，后端可能有人做Java）。所以，Java也被淘汰出局。

- 如果选择Python。首先，Python是跨平台的。其次，Python的学习速度很快，3~5年经验的人，上手时间顶多3~5天。再次，Python对于文件，文本，命令行处理，支持的非常之好。最后，Python里面也有方便的图形化工具，例如Qt就提供了Python版本。

所以，选择Python。

**Part2.数组的作用**
============

如果没有数组，那么在遇到成序列的数据时候，比如设计怪物AI中的技能部分，表的结构就会是类似这个样子：
<table>
    <tbody>
        <tr>
            <th>
                length
            </th>
            <th>
                skill1
            </th>
            <th>
                <span style="font-weight:bold;text-align:center;background-color:#f7f7f7;">skill2</span>
            </th>
            <th>
                <span style="font-weight:bold;text-align:center;background-color:#f7f7f7;">skill3</span>
            </th>
            <th>
                <span style="font-weight:bold;text-align:center;background-color:#f7f7f7;">skill4</span>
            </th>
        </tr>
        <tr>
            <td>
                4
            </td>
            <td>
                火球
            </td>
            <td>
                冰箭
            </td>
            <td>
                魔法盾
            </td>
            <td>
                顺移
            </td>
        </tr>
        <tr>
            <td>
                3
            </td>
            <td>
                突刺
            </td>
            <td>
                半月
            </td>
            <td>
                重斩
            </td>
            <td>
                <br />
            </td>
        </tr>
        <tr>
            <td>
                1
            </td>
            <td>
                治疗
            </td>
            <td>
                <br />
            </td>
            <td>
                <br />
            </td>
            <td>
                <br />
            </td>
        </tr>
    </tbody>
</table>



如果你使用过类似这样的JSON结构，那么你应该知道，在填写数据的时候，容易出错，输出数据的时候会很难看（不管是填充null作为空数据，还是不输出空白格，都一样难看。前者存在无用数据，后者丢失了表的结构，造成阅读困难），遍历代码写起来也很麻烦。

在JSON中，数组天生就可以获得其“元素个数”，并且可以方便的遍历。所以，我们要在工具层面支持数组，这样才能使用JSON的这个特性。

**Part3.“引用”该怎么用？**
============

还是举一个例子，在经营建造游戏中，对于建筑物属性的定义，每个建筑的解锁等级这是一个固定值，该建筑占用的地块面积也是一个固定值。但是该建筑不同等级的属性，则是完全不相同的。如果是一个资源产生建筑，那么会有不同的资源生成速度和资源上限，如果是一个出兵建筑，会有可造兵种类别，出兵时间。如果是一个防御建筑，会有攻击半径，伤害力等。这些不同结构的字段，是没有可能放到一个二维表中的。

一般采用的方式是，会有几种方案：
1.会有一个主要的表来存放所有建筑包含的相同的字段，然后那些不相同的字段信息放到其他表中，然后通过主键跳转来访问。
2.直接拆成多个表来填数据
3.使用一些不同的开关字段或分类字段，让同一个字段在不同开关状态下有不同的含义。现在游戏越来越复杂，这是最不建议的一种方式。

上面的3种方案，维护和修改成本都很高。

采用引用实现就很简单，还是多个表，然后在主要表上，插入其他表的引用即可。
<table>
    <tbody>
        <tr>
            <th>
                s
            </th>
            <th>
                i
            </th>
            <th>
                i
            </th>
            <th>
                r
            </th>
            <th>
                r
            </th>
            <th>
                r
            </th>
        </tr>
        <tr>
            <td>
                name
            </td>
            <td>
                unlock_lv
            </td>
            <td>
                area
            </td>
            <td>
                lv1
            </td>
            <td>
                lv2
            </td>
            <td>
                lv3
            </td>
        </tr>
        <tr>
            <td>
                基地
            </td>
            <td>
                1
            </td>
            <td>
                4
            </td>
            <td>
                基地.lv1
            </td>
            <td>
                基地.lv2
            </td>
            <td>
                基地.lv3
            </td>
        </tr>
        <tr>
            <td>
                铀矿
            </td>
            <td>
                3
            </td>
            <td>
                4
            </td>
            <td>
                铀矿.lv1
            </td>
            <td>
                铀矿.lv2
            </td>
            <td>
                铀矿.lv3
            </td>
        </tr>
        <tr>
            <td>
                兵营
            </td>
            <td>
                5
            </td>
            <td>
                1
            </td>
            <td>
                兵营.lv1
            </td>
            <td>
                兵营.lv2
            </td>
            <td>
                兵营.lv3
            </td>
        </tr>
    </tbody>
</table>

**Part4.主表模式的意义是什么？**
============

游戏开发中，前后端对于数据的需求是不一样的。前端需要的是一些显示数据，如资源名称，动作参数。后端需要的是一些计算数据，比如攻击力，防御力，伤害公式等。但是有一些数据，是前后端都需要的，比如：技能范围，技能类型等，这些数据既与前端的显示有关系也和后端的逻辑计算有关系。

那么这种情况下，按照传统方式，也会拆成若干表。一般是一张表前端用，一张表后端用。但问题在于，前后端都需要的数据该如何处理？在两个表之间同步是一个成本比较高的办法。

这就体现出主表模式的意义了。我们可以把这些数据都组织在一张表上：
<table style="border-collapse:collapse;border-spacing:0;">
    <tbody>
        <tr>
            <th style="font-family:arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;">
                name
            </th>
            <th style="font-family:arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;">
                type
            </th>
            <th style="font-family:arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;">
                effect
            </th>
            <th style="font-family:arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;">
                atk
            </th>
        </tr>
        <tr>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                平砍
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                1
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                平砍.png
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                10
            </td>
        </tr>
        <tr>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                横扫千军
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                3
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                横扫千军.png
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                7
            </td>
        </tr>
        <tr>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                暴风雪
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                4
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                暴风雪.png
            </td>
            <td style="font-family:arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;">
                8
            </td>
        </tr>
    </tbody>
</table>

然后在输出的时候，在主表模式中，分成两个来输出：
<table>
    <tbody>
        <tr>
            <th>
                skill-&gt;skill_fn
            </th>
            <th>
                name
            </th>
            <th>
                type
            </th>
            <th>
                effect
            </th>
        </tr>
        <tr>
            <td>
                skill-&gt;skill_bn
            </td>
            <td>
                name
            </td>
            <td>
                type
            </td>
            <td>
                atk
            </td>
        </tr>
    </tbody>
</table>
**Finally**
============

需求一直在变，工具要提供的是应对不同需求的灵活性。
