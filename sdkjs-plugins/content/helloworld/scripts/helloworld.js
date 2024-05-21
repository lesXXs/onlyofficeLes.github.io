/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

// // Example insert text into editors (different implementations)
// (function(window, undefined){
//
//     var text = "Hello world!";
//
//     window.Asc.plugin.init = function()
//     {
//         var variant = 2;
//
//         switch (variant)
//         {
//             case 0:
//             {
//                 // serialize command as text
//                 var sScript = "var oDocument = Api.GetDocument();";
//                 sScript += "oParagraph = Api.CreateParagraph();";
//                 sScript += "oParagraph.AddText('Hello world!');";
//                 sScript += "oDocument.InsertContent([oParagraph]);";
//                 this.info.recalculate = true;
//                 this.executeCommand("close", sScript);
//                 break;
//             }
//             case 1:
//             {
//                 // call command without external variables
//                 this.callCommand(function() {
//                     var oDocument = Api.GetDocument();
//                     var oParagraph = Api.CreateParagraph();
//                     oParagraph.AddText("Hello world!");
//                     oDocument.InsertContent([oParagraph]);
//                 }, true);
//                 break;
//             }
//             case 2:
//             {
//                 // call command with external variables
//                 Asc.scope.text = text; // export variable to plugin scope
//                 this.callCommand(function() {
//                     var oDocument = Api.GetDocument();
//                     var oParagraph = Api.CreateParagraph();
//                     oParagraph.AddText(Asc.scope.text); // or oParagraph.AddText(scope.text);
//                     oDocument.InsertContent([oParagraph]);
//                 }, true);
//                 break;
//             }
//             default:
//                 break;
//         }
//     };
//
//     window.Asc.plugin.button = function(id)
//     {
//     };
//
// })(window, undefined);
(function(window, undefined) {
    // 初始化插件
    window.Asc.plugin.init = function() {
        // 处理文档内容已准备就绪事件
        function handleContentReady() {
            // 获取文档内容
            var documentContent = Api.GetDocumentContent();

            // 提取新增文本
            var newText = extractNewText(documentContent);

            // TODO: 根据实际需求，实现对新增文本的关键实体和知识要素识别逻辑
            var extractionResult = extractKeyEntitiesAndKnowledgeElements(newText);
            var entities = extractionResult.entities;
            var knowledgeElements = extractionResult.knowledgeElements;

            // 高亮文本实体并在实体下方展示框
            highlightEntitiesAndShowContent(entities, knowledgeElements);
        }

        // 注册文档内容已准备就绪事件监听器
        this.attachEvent("onDocumentContentReady", handleContentReady);
    };

    window.Asc.plugin.button = function(id) {
        // 处理按钮点击事件（如果有需要）
    };

    // 提取新增文本的函数
    function extractNewText(documentContent) {
        // TODO: 根据实际需求，实现从文档内容中提取新增文本的逻辑
        // 您可以将当前文档内容与之前存储的内容进行比较，以确定新增的文本部分

        // 为了演示目的，我们假设整个文档内容都被视为新增文本
        return documentContent;
    }

    // 高亮文本实体并在实体下方展示框
    function highlightEntitiesAndShowContent(entities, knowledgeElements) {
        var oDocument = Api.GetDocument();

        // 遍历所有实体
        for (var i = 0; i < entities.length; i++) {
            var entity = entities[i];

            // 创建一个文本范围对象来标记实体位置
            var range = oDocument.CreateRange(entity.startIndex, entity.endIndex);

            // 设置高亮样式
            range.SetHighlightColor("#FFFF00"); // 设置为黄色高亮

            // 在实体下方创建一个框
            var oFrame = Api.CreateFrame();
            oFrame.SetWidth("100px"); // 设置框的宽度
            oFrame.SetHeight("50px"); // 设置框的高度
            oFrame.SetPosition(entity.startIndex, entity.endIndex + 1); // 设置框的位置，+1 是为了放在实体下方

            // 创建框内的内容
            var oParagraph = Api.CreateParagraph();
            oParagraph.AddText(knowledgeElements[i].description);
            oFrame.InsertContent([oParagraph]);

            // 将范围和框插入文档
            oDocument.InsertContent([range, oFrame]);
        }
    }
})(window, undefined);