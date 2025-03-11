<template>
  <div class="editor-header">
    <div class="left">
      <Popover trigger="click" placement="bottom-start" v-model:value="mainMenuVisible">
        <template #content>
          <PopoverMenuItem @click="openAIPPTDialog(); mainMenuVisible = false">AI 生成 PPT（测试版）</PopoverMenuItem>
          <FileInput accept="application/vnd.openxmlformats-officedocument.presentationml.presentation"  @change="files => {
            importPPTXFile(files)
            mainMenuVisible = false
          }">
            <PopoverMenuItem>导入 pptx 文件（测试版）</PopoverMenuItem>
          </FileInput>
          <FileInput accept=".pptist"  @change="files => {
            importSpecificFile(files)
            mainMenuVisible = false
          }">
            <PopoverMenuItem>导入 pptist 文件</PopoverMenuItem>
          </FileInput>
          <FileInput accept=".json"  @change="files => {
            importJSONFile(files)
            mainMenuVisible = false
          }">
            <PopoverMenuItem>导入 JSON 文件</PopoverMenuItem>
          </FileInput>
          <PopoverMenuItem @click="setDialogForExport('pptx')">导出文件</PopoverMenuItem>
          <PopoverMenuItem @click="resetSlides(); mainMenuVisible = false">重置幻灯片</PopoverMenuItem>
          <PopoverMenuItem @click="openMarkupPanel(); mainMenuVisible = false">幻灯片类型标注</PopoverMenuItem>
          <PopoverMenuItem @click="goLink('https://github.com/pipipi-pikachu/PPTist/issues')">意见反馈</PopoverMenuItem>
          <PopoverMenuItem @click="goLink('https://github.com/pipipi-pikachu/PPTist/blob/master/doc/Q&A.md')">常见问题</PopoverMenuItem>
          <PopoverMenuItem @click="mainMenuVisible = false; hotkeyDrawerVisible = true">快捷操作</PopoverMenuItem>
        </template>
        <div class="menu-item"><IconHamburgerButton class="icon" /></div>
      </Popover>

      <div class="title">
        <Input 
          class="title-input" 
          ref="titleInputRef"
          v-model:value="titleValue" 
          @blur="handleUpdateTitle()" 
          v-if="editingTitle" 
        ></Input>
        <div 
          class="title-text"
          @click="startEditTitle()"
          :title="title"
          v-else
        >{{ title }}</div>
      </div>
    </div>

    <div class="right">
      <div class="group-menu-item">
        <div class="menu-item" v-tooltip="'幻灯片放映（F5）'" @click="enterScreening()">
          <IconPpt class="icon" />
        </div>
        <Popover trigger="click" center>
          <template #content>
            <PopoverMenuItem @click="enterScreeningFromStart()">从头开始</PopoverMenuItem>
            <PopoverMenuItem @click="enterScreening()">从当前页开始</PopoverMenuItem>
          </template>
          <div class="arrow-btn"><IconDown class="arrow" /></div>
        </Popover>
      </div>
      <div class="menu-item" v-tooltip="'AI生成PPT'" @click="openAIPPTDialog(); mainMenuVisible = false">
        <span class="text ai">AI</span>
      </div>
      <div class="menu-item" v-tooltip="'导出'" @click="setDialogForExport('pptx')">
        <IconDownload class="icon" />
      </div>
      <a class="github-link" v-tooltip="'Copyright © 2020-PRESENT pipipi-pikachu'" href="https://github.com/pipipi-pikachu/PPTist" target="_blank">
        <div class="menu-item"><IconGithub class="icon" /></div>
      </a>
    </div>

    <Drawer
      :width="320"
      v-model:visible="hotkeyDrawerVisible"
      placement="right"
    >
      <HotkeyDoc />
      <template v-slot:title>快捷操作</template>
    </Drawer>

    <FullscreenSpin :loading="exporting" tip="正在导入..." />
  </div>
</template>

<script lang="ts" setup>
import { nextTick, ref } from 'vue'
import { storeToRefs } from 'pinia'
import { useMainStore, useSlidesStore } from '@/store'
import useScreening from '@/hooks/useScreening'
import useImport from '@/hooks/useImport'
import useSlideHandler from '@/hooks/useSlideHandler'
import type { DialogForExportTypes } from '@/types/export'
import message from '@/utils/message'
import { nanoid } from 'nanoid'

import HotkeyDoc from './HotkeyDoc.vue'
import FileInput from '@/components/FileInput.vue'
import FullscreenSpin from '@/components/FullscreenSpin.vue'
import Drawer from '@/components/Drawer.vue'
import Input from '@/components/Input.vue'
import Popover from '@/components/Popover.vue'
import PopoverMenuItem from '@/components/PopoverMenuItem.vue'
import type { Slide } from '@/types/slides'
import type { PPTElement, ShapeElement } from '@/types/element'

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const { title } = storeToRefs(slidesStore)
const { enterScreening, enterScreeningFromStart } = useScreening()
const { importSpecificFile, importPPTXFile, exporting } = useImport()
const { resetSlides } = useSlideHandler()
const { setSlides, setViewportRatio } = slidesStore

const mainMenuVisible = ref(false)
const hotkeyDrawerVisible = ref(false)
const editingTitle = ref(false)
const titleInputRef = ref<InstanceType<typeof Input>>()
const titleValue = ref('')

const startEditTitle = () => {
  titleValue.value = title.value
  editingTitle.value = true
  nextTick(() => titleInputRef.value?.focus())
}

const handleUpdateTitle = () => {
  slidesStore.setTitle(titleValue.value)
  editingTitle.value = false
}

const goLink = (url: string) => {
  window.open(url)
  mainMenuVisible.value = false
}

const setDialogForExport = (type: DialogForExportTypes) => {
  mainStore.setDialogForExport(type)
  mainMenuVisible.value = false
}

const openMarkupPanel = () => {
  mainStore.setMarkupPanelState(true)
}

const openAIPPTDialog = () => {
  mainStore.setAIPPTDialogState(true)
}

interface ImportedJSON {
  title?: string
  theme?: SlideTheme
  slides: Slide[]
  width?: number
  height?: number
}

interface ViewportCalculationResult {
  processedSlidesWithScale: Slide[]
  viewportRatio: number
}

const calculateViewportAndScale = (slides: Slide[], originalWidth?: number, originalHeight?: number): ViewportCalculationResult => {
  // 如果提供了原始尺寸，直接使用原始尺寸和比例
  if (originalWidth && originalHeight) {
    return {
      processedSlidesWithScale: slides,  // 直接使用原始slides，不进行缩放
      viewportRatio: originalHeight / originalWidth
    }
  }

  // 如果没有原始尺寸，使用默认值
  const defaultViewportSize = 1000
  const defaultRatio = 0.5625 // 16:9

  return {
    processedSlidesWithScale: slides,
    viewportRatio: defaultRatio
  }
}

const importJSONFile = (files: FileList) => {
  const file = files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = e => {
    try {
      const json = JSON.parse(e.target?.result as string) as ImportedJSON
      if (!json || !json.slides || !Array.isArray(json.slides)) {
        throw new Error('无效的 JSON 格式')
      }

      // 设置主题（如果存在）
      if (json.theme) {
        slidesStore.setTheme(json.theme)
      }

      // 设置标题（如果存在）
      if (json.title) {
        slidesStore.setTitle(json.title)
      }

      // 设置视口大小（如果存在）
      if (json.width) {
        slidesStore.setViewportSize(json.width)
      }

      // 处理每个幻灯片并计算合适的视口比例
      const { processedSlidesWithScale, viewportRatio } = calculateViewportAndScale(
        json.slides,
        json.width,
        json.height
      )

      // 设置视口比例
      slidesStore.setViewportRatio(viewportRatio)

      // 处理每个幻灯片的基本属性和样式
      const processedSlides = processedSlidesWithScale.map(slide => {
        // 确保每个幻灯片有唯一ID，保持原始背景设置
        const processedSlide = {
          ...slide,
          id: slide.id || nanoid(10)
        }

        // 如果没有背景设置，使用默认背景
        if (!processedSlide.background) {
          processedSlide.background = {
            type: 'solid',
            color: '#000000'
          }
        }

        // 处理元素
        if (processedSlide.elements) {
          processedSlide.elements = processedSlide.elements.map(element => {
            const processedElement = {
              ...element,
              id: element.id || nanoid(10)
            }

            // 处理文本元素
            if (processedElement.type === 'text') {
              if (!processedElement.defaultColor) {
                processedElement.defaultColor = json.theme?.fontColor || '#333'
              }
              if (!processedElement.defaultFontName) {
                processedElement.defaultFontName = json.theme?.fontName || ''
              }
            }

            // 处理形状元素
            if (processedElement.type === 'shape') {
              // 保持原始fill属性
              if ('fill' in processedElement) {
                const fill = (processedElement as any).fill
                if (fill) {
                  // 保持所有颜色格式
                  if (typeof fill === 'string' && (
                    fill.startsWith('rgb') || 
                    fill.startsWith('rgba') || 
                    fill.startsWith('#')
                  )) {
                    processedElement.fill = fill
                  }
                  // 其他格式转换为background
                  else {
                    processedElement.background = typeof fill === 'string'
                      ? { type: 'solid', color: fill }
                      : fill
                    delete (processedElement as any).fill
                  }
                }
              }
              
              // 如果有background属性，确保颜色格式正确
              if (processedElement.background && processedElement.background.type === 'solid') {
                const color = processedElement.background.color
                if (typeof color === 'string' && (
                  color.startsWith('rgb') || 
                  color.startsWith('rgba') || 
                  color.startsWith('#')
                )) {
                  processedElement.background.color = color
                }
              }

              // 确保有viewBox属性
              if (!processedElement.viewBox && processedElement.width && processedElement.height) {
                processedElement.viewBox = [processedElement.width, processedElement.height]
              }
            }

            // 处理线条元素
            if (processedElement.type === 'line') {
              if (!processedElement.color) {
                processedElement.color = json.theme?.themeColor || '#333'
              }
              // 保持原始颜色格式
              else if (typeof processedElement.color === 'string' && (
                processedElement.color.startsWith('rgb') || 
                processedElement.color.startsWith('rgba') || 
                processedElement.color.startsWith('#')
              )) {
                processedElement.color = processedElement.color
              }
            }

            // 确保基本属性存在
            if (typeof processedElement.left !== 'number') processedElement.left = 0
            if (typeof processedElement.top !== 'number') processedElement.top = 0
            if (typeof processedElement.rotate !== 'number') processedElement.rotate = 0
            if (typeof processedElement.opacity !== 'number') processedElement.opacity = 1

            // 处理阴影颜色
            if (processedElement.shadow && typeof processedElement.shadow.color === 'string') {
              const shadowColor = processedElement.shadow.color
              if (shadowColor.startsWith('rgb') || shadowColor.startsWith('rgba') || shadowColor.startsWith('#')) {
                processedElement.shadow.color = shadowColor
              }
            }

            return processedElement
          })
        }

        return processedSlide
      })

      // 设置幻灯片
      slidesStore.setSlides(processedSlides)

      message.success('导入成功')
    } catch (error) {
      console.error('Import error:', error)
      message.error('导入失败：' + (error.message || '未知错误'))
    }
  }
  reader.readAsText(file)
}
</script>

<style lang="scss" scoped>
.editor-header {
  background-color: #fff;
  user-select: none;
  border-bottom: 1px solid $borderColor;
  display: flex;
  justify-content: space-between;
  padding: 0 5px;
}
.left, .right {
  display: flex;
  justify-content: center;
  align-items: center;
}
.menu-item {
  height: 30px;
  display: flex;
  justify-content: center;
  align-items: center;
  font-size: 14px;
  padding: 0 10px;
  border-radius: $borderRadius;
  cursor: pointer;

  .icon {
    font-size: 18px;
    color: #666;
  }
  .text {
    width: 18px;
    text-align: center;
    font-size: 17px;
  }
  .ai {
    background: linear-gradient(270deg, #d897fd, #33bcfc);
    background-clip: text;
    color: transparent;
    font-weight: 700;
  }

  &:hover {
    background-color: #f1f1f1;
  }
}
.group-menu-item {
  height: 30px;
  display: flex;
  margin: 0 8px;
  padding: 0 2px;
  border-radius: $borderRadius;

  &:hover {
    background-color: #f1f1f1;
  }

  .menu-item {
    padding: 0 3px;
  }
  .arrow-btn {
    display: flex;
    justify-content: center;
    align-items: center;
    cursor: pointer;
  }
}
.title {
  height: 30px;
  margin-left: 2px;
  font-size: 13px;

  .title-input {
    width: 200px;
    height: 100%;
    padding-left: 0;
    padding-right: 0;

    ::v-deep(input) {
      height: 28px;
      line-height: 28px;
    }
  }
  .title-text {
    min-width: 20px;
    max-width: 400px;
    line-height: 30px;
    padding: 0 6px;
    border-radius: $borderRadius;
    cursor: pointer;

    @include ellipsis-oneline();

    &:hover {
      background-color: #f1f1f1;
    }
  }
}
.github-link {
  display: inline-block;
  height: 30px;
}
</style>